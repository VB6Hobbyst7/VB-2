VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Interface Program"
   ClientHeight    =   10665
   ClientLeft      =   1065
   ClientTop       =   750
   ClientWidth     =   16260
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
   ScaleHeight     =   10665
   ScaleWidth      =   16260
   Begin FPSpread.vaSpread vasTransTemp 
      Height          =   2865
      Left            =   11220
      TabIndex        =   75
      Top             =   5250
      Visible         =   0   'False
      Width           =   4605
      _Version        =   393216
      _ExtentX        =   8123
      _ExtentY        =   5054
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
      SpreadDesigner  =   "frmInterface.frx":0442
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   405
      Left            =   18480
      TabIndex        =   49
      Top             =   240
      Width           =   1275
   End
   Begin VB.Frame Frame6 
      Caption         =   "BARCODE ERR"
      Height          =   9225
      Left            =   16230
      TabIndex        =   46
      Top             =   900
      Visible         =   0   'False
      Width           =   2565
      Begin FPSpread.vaSpread vasBARERR 
         Height          =   8235
         Left            =   90
         TabIndex        =   48
         Top             =   780
         Width           =   2355
         _Version        =   393216
         _ExtentX        =   4154
         _ExtentY        =   14526
         _StockProps     =   64
         ColHeaderDisplay=   0
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
         MaxCols         =   1
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "frmInterface.frx":06CC
      End
      Begin VB.CommandButton cmdLBCLEAR 
         Caption         =   "목록정리"
         Height          =   435
         Left            =   90
         TabIndex        =   47
         Top             =   270
         Width           =   2355
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9915
      Left            =   60
      TabIndex        =   10
      Top             =   750
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   17489
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Interface"
      TabPicture(0)   =   "frmInterface.frx":1F31
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":1F4D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdAbnormal"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdAbnormal 
         Caption         =   "이상검체 조회"
         Height          =   405
         Left            =   -62730
         TabIndex        =   83
         Top             =   8220
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame Frame1 
         Height          =   9450
         Left            =   -74850
         TabIndex        =   23
         Top             =   360
         Width           =   15810
         Begin VB.Frame fraLogin 
            Height          =   765
            Left            =   11310
            TabIndex        =   104
            Top             =   450
            Width           =   4335
            Begin VB.CommandButton cmdLogin 
               BackColor       =   &H0000C000&
               Caption         =   "로그인"
               Height          =   495
               Left            =   90
               Style           =   1  '그래픽
               TabIndex        =   105
               Top             =   150
               Width           =   1305
            End
            Begin Threed.SSPanel sspUsername 
               Height          =   495
               Left            =   1470
               TabIndex        =   106
               Top             =   150
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   873
               _Version        =   131074
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   14.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label lblSendID 
               Height          =   105
               Left            =   2640
               TabIndex        =   107
               Top             =   690
               Visible         =   0   'False
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdResSave 
            Caption         =   "결과저장"
            Height          =   465
            Left            =   14430
            TabIndex        =   99
            Top             =   1410
            Width           =   1305
         End
         Begin VB.Frame Frame8 
            Caption         =   "hidden"
            Height          =   1725
            Left            =   1320
            TabIndex        =   86
            Top             =   4560
            Visible         =   0   'False
            Width           =   7695
            Begin VB.CommandButton cmdAddDataCall 
               Caption         =   "데이터 불러오기"
               Height          =   405
               Left            =   5130
               TabIndex        =   88
               Top             =   1320
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.CommandButton Command8 
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
               Left            =   6990
               TabIndex        =   87
               Top             =   1320
               Visible         =   0   'False
               Width           =   1905
            End
            Begin MSComCtl2.DTPicker dtpAddExamDate1 
               Height          =   315
               Left            =   1410
               TabIndex        =   89
               Top             =   1350
               Visible         =   0   'False
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   393216
               Format          =   103481345
               CurrentDate     =   40780
            End
            Begin MSComCtl2.DTPicker dtpAddExamDate2 
               Height          =   315
               Left            =   3210
               TabIndex        =   90
               Top             =   1350
               Visible         =   0   'False
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   556
               _Version        =   393216
               Format          =   103481345
               CurrentDate     =   40780
            End
            Begin VB.Label Label7 
               Caption         =   "추가항목 조회"
               Height          =   405
               Left            =   450
               TabIndex        =   93
               Top             =   1050
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label Label8 
               Caption         =   "검사일자"
               Height          =   225
               Left            =   510
               TabIndex        =   92
               Top             =   1410
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label Label9 
               Caption         =   "-"
               Height          =   225
               Left            =   3000
               TabIndex        =   91
               Top             =   1410
               Visible         =   0   'False
               Width           =   135
            End
         End
         Begin FPSpread.vaSpread vasLISTRES 
            Height          =   7395
            Left            =   11220
            TabIndex        =   61
            Top             =   1950
            Width           =   4515
            _Version        =   393216
            _ExtentX        =   7964
            _ExtentY        =   13044
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            EditModeReplace =   -1  'True
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
            MaxCols         =   13
            MaxRows         =   100
            Protect         =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":1F69
         End
         Begin VB.CommandButton cmdVasListWidth 
            Caption         =   ">>"
            Height          =   405
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.CheckBox ChkAll 
            Height          =   285
            Left            =   570
            TabIndex        =   33
            Top             =   1500
            Width           =   225
         End
         Begin VB.Frame Frame4 
            Caption         =   "[검사결과조회]"
            Height          =   1125
            Left            =   90
            TabIndex        =   24
            Top             =   240
            Width           =   15645
            Begin VB.TextBox txtWorkNo 
               Height          =   315
               Left            =   6180
               TabIndex        =   102
               Top             =   720
               Width           =   1785
            End
            Begin VB.CheckBox cnkNEMD 
               Caption         =   "신장내과"
               Height          =   255
               Left            =   6210
               TabIndex        =   98
               Top             =   450
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CheckBox chkMicro 
               Caption         =   "검경대상"
               Height          =   255
               Left            =   6210
               TabIndex        =   97
               Top             =   150
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.ComboBox cboMicro 
               Height          =   315
               ItemData        =   "frmInterface.frx":2C3C
               Left            =   3150
               List            =   "frmInterface.frx":2C3E
               TabIndex        =   95
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtBarcode 
               Height          =   315
               Left            =   9390
               TabIndex        =   29
               Top             =   720
               Width           =   1785
            End
            Begin VB.ComboBox cmbTransGubun 
               Height          =   315
               ItemData        =   "frmInterface.frx":2C40
               Left            =   990
               List            =   "frmInterface.frx":2C50
               TabIndex        =   28
               Text            =   "전체"
               Top             =   720
               Width           =   1455
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
               Left            =   6780
               TabIndex        =   27
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
               Left            =   8610
               TabIndex        =   26
               Top             =   210
               Width           =   1275
            End
            Begin VB.CommandButton cmdListTrans 
               BackColor       =   &H0000FF00&
               Caption         =   "결과전송"
               Enabled         =   0   'False
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
               Left            =   9930
               TabIndex        =   25
               Top             =   210
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker dtpExamDate 
               Height          =   315
               Left            =   990
               TabIndex        =   30
               Top             =   270
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   103481345
               CurrentDate     =   40780
            End
            Begin MSComCtl2.DTPicker dtpExamDate2 
               Height          =   315
               Left            =   2670
               TabIndex        =   85
               Top             =   270
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   556
               _Version        =   393216
               Format          =   103481345
               CurrentDate     =   40780
            End
            Begin VB.Label Label13 
               Caption         =   "접수번호 검색"
               Height          =   225
               Left            =   4770
               TabIndex        =   103
               Top             =   810
               Width           =   1395
            End
            Begin VB.Label Label12 
               Caption         =   "Micro"
               Height          =   225
               Left            =   2580
               TabIndex        =   96
               Top             =   780
               Width           =   585
            End
            Begin VB.Label Label11 
               Caption         =   "-"
               Height          =   225
               Left            =   2490
               TabIndex        =   84
               Top             =   330
               Width           =   135
            End
            Begin VB.Label Label2 
               Caption         =   "검사일자"
               Height          =   225
               Left            =   90
               TabIndex        =   60
               Top             =   330
               Width           =   915
            End
            Begin VB.Label Label4 
               Caption         =   "Barcode 검색"
               Height          =   225
               Left            =   8100
               TabIndex        =   32
               Top             =   780
               Width           =   1275
            End
            Begin VB.Label Label3 
               Caption         =   "전송구분"
               Height          =   225
               Left            =   90
               TabIndex        =   31
               Top             =   780
               Width           =   885
            End
         End
         Begin FPSpread.vaSpread vasList 
            Height          =   7935
            Left            =   90
            TabIndex        =   81
            Top             =   1410
            Width           =   10995
            _Version        =   393216
            _ExtentX        =   19394
            _ExtentY        =   13996
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
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
            MaxCols         =   107
            Protect         =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":2C6E
         End
         Begin VB.Label lblInfo2 
            AutoSize        =   -1  'True
            Caption         =   "Barcode :"
            Height          =   195
            Left            =   11220
            TabIndex        =   101
            Top             =   1710
            Width           =   945
         End
         Begin VB.Label lblInfo1 
            AutoSize        =   -1  'True
            Caption         =   "Barcode :"
            Height          =   195
            Left            =   11220
            TabIndex        =   100
            Top             =   1410
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Height          =   9420
         Left            =   120
         TabIndex        =   16
         Top             =   390
         Width           =   15900
         Begin FPSpread.vaSpread vasResTemp 
            Height          =   2355
            Left            =   1410
            TabIndex        =   82
            Top             =   6000
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
            SpreadDesigner  =   "frmInterface.frx":AB17
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Command10"
            Height          =   405
            Left            =   1500
            TabIndex        =   80
            Top             =   1560
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Frame Frame7 
            Caption         =   "[Manual Order]"
            Height          =   585
            Left            =   3930
            TabIndex        =   76
            Top             =   120
            Visible         =   0   'False
            Width           =   5505
            Begin VB.OptionButton optOrder 
               Caption         =   "Cancel"
               Height          =   315
               Index           =   3
               Left            =   4380
               TabIndex        =   94
               Top             =   210
               Width           =   1035
            End
            Begin VB.OptionButton optOrder 
               Caption         =   "Urine"
               Height          =   315
               Index           =   0
               Left            =   210
               TabIndex        =   79
               Top             =   210
               Width           =   1035
            End
            Begin VB.OptionButton optOrder 
               Caption         =   "Sieve"
               Height          =   315
               Index           =   1
               Left            =   1410
               TabIndex        =   78
               Top             =   210
               Width           =   1035
            End
            Begin VB.OptionButton optOrder 
               Caption         =   "Urine + Micro"
               Height          =   315
               Index           =   2
               Left            =   2670
               TabIndex        =   77
               Top             =   210
               Value           =   -1  'True
               Width           =   1815
            End
         End
         Begin FPSpread.vaSpread vasOrderTest 
            Height          =   4425
            Left            =   9480
            TabIndex        =   73
            Top             =   2520
            Visible         =   0   'False
            Width           =   4995
            _Version        =   393216
            _ExtentX        =   8811
            _ExtentY        =   7805
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
            SpreadDesigner  =   "frmInterface.frx":ADA1
         End
         Begin VB.Frame Frame5 
            Caption         =   "Frame5"
            Height          =   5865
            Left            =   1080
            TabIndex        =   43
            Top             =   1980
            Visible         =   0   'False
            Width           =   13065
            Begin VB.CommandButton Command9 
               Caption         =   "TEST"
               Height          =   1185
               Left            =   5820
               TabIndex        =   74
               Top             =   1920
               Width           =   1605
            End
            Begin VB.CommandButton cmdBARERR 
               Caption         =   "BARCODE ERR"
               Height          =   615
               Left            =   7890
               TabIndex        =   64
               Top             =   4350
               Visible         =   0   'False
               Width           =   2175
            End
            Begin FPSpread.vaSpread vasOrder 
               Height          =   1425
               Left            =   60
               TabIndex        =   57
               Top             =   4020
               Width           =   7335
               _Version        =   393216
               _ExtentX        =   12938
               _ExtentY        =   2514
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
               SpreadDesigner  =   "frmInterface.frx":B02B
            End
            Begin FPSpread.vaSpread vasConvert 
               Height          =   1665
               Left            =   150
               TabIndex        =   59
               Top             =   4050
               Width           =   6705
               _Version        =   393216
               _ExtentX        =   11827
               _ExtentY        =   2937
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
               SpreadDesigner  =   "frmInterface.frx":F537
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Command7"
               Height          =   585
               Left            =   5820
               TabIndex        =   56
               Top             =   330
               Width           =   1605
            End
            Begin VB.TextBox txtWinSockBuff 
               Height          =   3165
               Left            =   7650
               MultiLine       =   -1  'True
               TabIndex        =   55
               Top             =   450
               Width           =   5175
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Command6"
               Height          =   435
               Left            =   6030
               TabIndex        =   53
               Top             =   900
               Width           =   885
            End
            Begin VB.TextBox Text3 
               Height          =   855
               Left            =   120
               TabIndex        =   52
               Top             =   2460
               Width           =   2655
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Command5"
               Height          =   585
               Left            =   5790
               TabIndex        =   51
               Top             =   900
               Width           =   1215
            End
            Begin VB.TextBox Text2 
               Height          =   615
               Left            =   90
               TabIndex        =   50
               Top             =   3300
               Width           =   5535
            End
            Begin VB.TextBox Text1 
               Height          =   2265
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   45
               Top             =   270
               Width           =   5415
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Command1"
               Height          =   1155
               Left            =   5790
               TabIndex        =   44
               Top             =   270
               Width           =   1605
            End
            Begin Threed.SSPanel SSPanel3 
               Height          =   465
               Left            =   7980
               TabIndex        =   65
               Top             =   4410
               Visible         =   0   'False
               Width           =   3315
               _ExtentX        =   5847
               _ExtentY        =   820
               _Version        =   131074
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.CheckBox chkArchitect 
                  Caption         =   "Architect"
                  Height          =   345
                  Left            =   1740
                  TabIndex        =   67
                  Top             =   60
                  Width           =   1305
               End
               Begin VB.CheckBox chkCobas 
                  Caption         =   "Cobas"
                  Height          =   345
                  Left            =   240
                  TabIndex        =   66
                  Top             =   60
                  Value           =   1  '확인
                  Width           =   1305
               End
            End
            Begin Threed.SSPanel SSPanel2 
               Height          =   495
               Left            =   7890
               TabIndex        =   68
               Top             =   4380
               Visible         =   0   'False
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   873
               _Version        =   131074
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin MSComCtl2.DTPicker dtpEDate 
                  Height          =   345
                  Left            =   1860
                  TabIndex        =   69
                  Top             =   90
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   103481345
                  CurrentDate     =   41009
               End
               Begin MSComCtl2.DTPicker dtpSDate 
                  Height          =   345
                  Left            =   150
                  TabIndex        =   70
                  Top             =   90
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   103481345
                  CurrentDate     =   41009
               End
               Begin VB.Label Label5 
                  Caption         =   "-"
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   71
                  Top             =   120
                  Width           =   165
               End
            End
         End
         Begin VB.CommandButton cmdImage 
            Caption         =   "이미지 확인"
            Height          =   585
            Left            =   10560
            TabIndex        =   62
            Top             =   990
            Visible         =   0   'False
            Width           =   2535
         End
         Begin FPSpread.vaSpread vasListTemp 
            Height          =   1605
            Left            =   3660
            TabIndex        =   58
            Top             =   2340
            Visible         =   0   'False
            Width           =   9345
            _Version        =   393216
            _ExtentX        =   16484
            _ExtentY        =   2831
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
            SpreadDesigner  =   "frmInterface.frx":13A3B
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   525
            Left            =   9930
            TabIndex        =   42
            Top             =   2550
            Visible         =   0   'False
            Width           =   2775
         End
         Begin FPSpread.vaSpread vasTMaxList 
            Height          =   2805
            Left            =   3870
            TabIndex        =   40
            Top             =   3360
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
            SpreadDesigner  =   "frmInterface.frx":13CC5
         End
         Begin VB.CommandButton Command3 
            Caption         =   "QC 결과전송"
            Height          =   405
            Left            =   11070
            TabIndex        =   39
            Top             =   330
            Visible         =   0   'False
            Width           =   1695
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   1155
            Left            =   2730
            TabIndex        =   38
            Top             =   3690
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
            SpreadDesigner  =   "frmInterface.frx":13F4F
         End
         Begin VB.TextBox txtData 
            Height          =   1215
            Left            =   11580
            TabIndex        =   37
            Top             =   6600
            Visible         =   0   'False
            Width           =   2715
         End
         Begin FPSpread.vaSpread vasOrderBuf 
            Height          =   1215
            Left            =   7200
            TabIndex        =   36
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
            SpreadDesigner  =   "frmInterface.frx":141D9
         End
         Begin VB.CommandButton cmdVasIDWidth 
            Caption         =   ">>"
            Height          =   405
            Left            =   90
            TabIndex        =   34
            Top             =   780
            Width           =   405
         End
         Begin VB.TextBox txtBuff 
            Height          =   1215
            Left            =   7920
            TabIndex        =   20
            Top             =   3120
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
            Left            =   540
            TabIndex        =   17
            Top             =   840
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   8565
            Left            =   60
            TabIndex        =   21
            Top             =   750
            Width           =   8385
            _Version        =   393216
            _ExtentX        =   14790
            _ExtentY        =   15108
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   11.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   16777215
            MaxCols         =   107
            Protect         =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":186E5
         End
         Begin FPSpread.vaSpread vasTMaxRes 
            Height          =   2235
            Left            =   7650
            TabIndex        =   41
            Top             =   5700
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
            SpreadDesigner  =   "frmInterface.frx":1C30A
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8175
            Left            =   8640
            TabIndex        =   22
            Top             =   750
            Width           =   7155
            _Version        =   393216
            _ExtentX        =   12621
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
            MaxCols         =   13
            MaxRows         =   100
            Protect         =   0   'False
            SpreadDesigner  =   "frmInterface.frx":1C594
         End
         Begin VB.Label Label10 
            Caption         =   "※이미지를 확인하시려면 더블클릭 해주세요"
            Height          =   225
            Left            =   3900
            TabIndex        =   72
            Top             =   360
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.Label lblRowNum 
            Caption         =   "Label10"
            Height          =   1095
            Left            =   7110
            TabIndex        =   63
            Top             =   1200
            Visible         =   0   'False
            Width           =   2745
         End
         Begin VB.Label Label6 
            Caption         =   "접속에러!!"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   13500
            TabIndex        =   54
            Top             =   330
            Visible         =   0   'False
            Width           =   1695
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
      _Version        =   131074
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
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   979
      _Version        =   131074
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
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   7230
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer Timer2 
         Interval        =   2000
         Left            =   5580
         Top             =   90
      End
      Begin VB.Timer Timer3 
         Interval        =   200
         Left            =   4950
         Top             =   60
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   3510
         Top             =   60
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   6360
         Top             =   60
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
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
         Picture         =   "frmInterface.frx":1D210
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   14
         Top             =   180
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   13050
         TabIndex        =   13
         Top             =   120
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103481344
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
         Left            =   12060
         TabIndex        =   12
         Top             =   180
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
         Left            =   5220
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
         SpreadDesigner  =   "frmInterface.frx":1D79A
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
         SpreadDesigner  =   "frmInterface.frx":21D0A
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
      Begin VB.Menu mnuOrder 
         Caption         =   "오더정보"
      End
      Begin VB.Menu mnuExamMst 
         Caption         =   "결과변환"
      End
      Begin VB.Menu mnuTransSet 
         Caption         =   "자동전송 설정"
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
Const colBarCode = 5
Const colRackPos = 3
Const colPID = 6
Const colPName = 7
Const colPSex = 8
Const colDept = 9
Const colWkNo = 4
Const colOrdStick = 10
Const colOrdMicro = 11
Const colOrdMicroYN = 12
Const colState = 13
Const colStickCheck = 14
Const colMicroCheck = 15
Const colImage = 15 '사용안함
Const colRStart = 16

' 장비코드 검사코드 검사명 수치결과 문자결과 seq
Const colEquipExam = 1
Const colExamCode = 2
Const colExamName = 3
Const colResValue = 4
Const colResult = 5
Const colSeq = 6
Const colResDate = 7
Const colResTime = 8

Const colAFLAG = 9
Const colDFLAG = 10
Const colPFLAG = 11
Const colRESFLAG = 12
Const colORDERFLAG = 13

Public gRow As Long
Dim sOrder As String
Dim ConfirmData As String
Dim gExamdate As String
Dim gExamTime As String
Dim sSampleType As String
Dim lsFlag As String
Dim llRow As Long

Dim blTimerChk As Boolean
Dim blTimerTerminate As Boolean
Dim gRecordCnt As Integer
Dim gPatCnt As Integer

Dim gOrderMsg(0 To 4) As String

Dim strQMode As String
Dim strTransYN As String '결과가 없는 데이터가 있을때 결과를 전송시키지 않게 하기 위해 만듬(EX: 볼륨에러)
Dim intTimer As Integer

Dim strFSend As String

Dim intErrorCheck As Integer


Dim strHeader As String
Dim strTerminate As String

'검사결과를 지울경우.
'오더헝태를 기억한다.
Dim strORDERType As String





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
Dim vasIDRow As Integer
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
    For vasIDRow = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = vasIDRow
        
        If vasID.Value = 1 Then '체크된 열은 저장이 안됨
'        If vasID.Value = "" Then
        
            liRet = -1
            
    
            liRet = Insert_Data(vasIDRow, vasID, Format(dtpToday, "YYYYMMDD"))
            
'            liRet = Insert_Data_SHINWON(vasIDRow, vasID, Format(dtpToday, "yyyymmdd"))
            
            If liRet = 1 Then
                'db_Commit gServer
                
                'SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "완료", vasIDRow, colState
            Else
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "실패", vasIDRow, colState
            End If
            vasID.Col = 1
            vasID.Row = vasIDRow
            vasID.Value = 0
        Else
        
        End If
    Next vasIDRow
    
End Sub

Function Insert_Data_SHINWON(argSpcRow As Integer, argSPR As vaSpread, argDate As String) As Integer
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
    
    Dim strAcptdd As String
    Dim strAcptno  As String
    
    Dim CmEx    As ADODB.Command
On Error GoTo errhadle

    Insert_Data_SHINWON = -1
    
    
    lsID = ""
    lsID = Trim(GetText(argSPR, argSpcRow, colBarCode))
    
    sTransDate = Format(GetDateFull, "yyyymmdd")
    sTransTime = Format(GetDateFull, "hhmmss")
    
    
    If IsNumeric(lsID) = False Or Len(lsID) <> 15 Then Exit Function
    
    strAcptdd = "20" & Mid(lsID, 1, 6)
    strAcptno = Mid(lsID, 7, 7)
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasTemp
    
'    SQL = " Select a.equipcode, a.examcode, a.resvalue, a.result, b.resgubun " & vbCrLf & _
'          " From pat_res a, equipexam b " & vbCrLf & _
'          " Where a.equipno = b.equipno " & vbCrLf & _
'          " And a.equipcode = b.equipcode " & vbCrLf & _
'          " And a.equipno = '" & gEquip & "' " & vbCrLf & _
'          " And a.barcode = '" & lsID & "' " & vbCrLf & _
'          " And a.EXAMDATE = '" & argDATE & "' "
    
    
    SQL = " Select a.equipcode, a.examcode, a.resvalue, a.result, b.resgubun, A.REFFLAG, A.PANICFLAG " & vbCrLf & _
          " From pat_res a, equipexam b " & vbCrLf & _
          " Where a.examcode = b.examcode " & vbCrLf & _
          " And a.equipcode = b.equipcode " & vbCrLf & _
          " And a.barcode = '" & lsID & "' " & vbCrLf & _
          " And a.EXAMDATE = '" & argDate & "' " & vbCrLf & _
          " AND a.resvalue <> ''"
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    
    Dim strWKSEQ    As String
    Dim strINT_SNO    As String


    SQL = ""
    SQL = SQL & vbCrLf & "SELECT MAX(WKSEQ) FROM SHINRSLT..INTRSLT "
    SQL = SQL & vbCrLf & " WHERE ACPTDD = '" & "20" & Mid(lsID, 1, 6) & "' "
    SQL = SQL & vbCrLf & "   AND acptno = '" & Mid(lsID, 7, 7) & "' "
    SQL = SQL & vbCrLf & "   AND WKDATE >= '" & "20" & Mid(lsID, 1, 6) & "' "
    
    res = db_select_Col(gServer, SQL)
    strWKSEQ = Val(gReadBuf(0)) + 1
    
    strINT_SNO = "0"
    sCnt = ""
    '서버로 결과값 저장하기

'    DisConnect_Server
'    Connect_Server
    
    Dim cntSEND As Integer
    cntSEND = 0
    
    For i = 1 To vasTemp.DataRowCnt
        cntSEND = cntSEND + 1
        
        If cntSEND = 5 Then
'            DisConnect_Server
'            Connect_Server
            cntSEND = 0
        End If
        
        sExamCode = Trim(GetText(vasTemp, i, 2))
        sResValue = Trim(GetText(vasTemp, i, 3))
        sResult = Trim(GetText(vasTemp, i, 4))
        sResGubun = Trim(GetText(vasTemp, i, 5))
        
        If sResGubun = "1" Then '문자
            sTransRes = sResult ' & "(" & sResValue & ")"
        Else
            sTransRes = sResValue
            sResult = ""
        End If
        
        
        If Trim(sExamCode) <> "" And Trim(sResValue) <> "" Then

'                           update shinrslt..testrslt
'                  set rslt      = '결과'
'                    , judg      = '판정'
'                    , panicflag = '판정플레그' --''
'                    , deltaflag = ''
'                    , rsltdt    = '일자시분초'
'                    , rgstr     = '인터페이스유저'
'                    , prcsflag  = '1' --고정
'                    , eqmt      = 'cobas 8000' --고정
'                    , cnfmdt    = '' -- 빈칸
'                    , cnfmstr   = ''-- 빈칸
'                    , cnfmflag  = ''-- 빈칸
'                    , crr  = 'CRR판정값'
'                    , ccr  = 'ccr판정값'
'                where acptdd    = '일자'
'                  and acptno    = '번호'
'                  and testcd_cd = '검사코드'
'                  and cnfmflag <> '1'--고정;
        
        
                           SQL = "update shinrslt..testrslt "
            SQL = SQL & vbCrLf & "   set rslt = '" & Trim(sTransRes) & "'"
            SQL = SQL & vbCrLf & "     , judg = '" & Trim(GetText(vasTemp, i, 6)) & "'"
            SQL = SQL & vbCrLf & "     , panicflag = '" & Trim(GetText(vasTemp, i, 7)) & "'"
            SQL = SQL & vbCrLf & "     , deltaflag = ''"
            SQL = SQL & vbCrLf & "     , rsltdt = '" & Trim(sTransDate) & Trim(sTransTime) & "'"
            SQL = SQL & vbCrLf & "     , rgstr = 'INTERFACE'"
            SQL = SQL & vbCrLf & "     , prcsflag = '1'"
            SQL = SQL & vbCrLf & "     , eqmt = 'cobas 8000'"
            SQL = SQL & vbCrLf & "     , cnfmdt = ''"
            SQL = SQL & vbCrLf & "     , cnfmstr = ''"
            SQL = SQL & vbCrLf & "     , cnfmflag = ''"
            SQL = SQL & vbCrLf & "     , crr = '' "
            SQL = SQL & vbCrLf & "     , ccr = '' "
            SQL = SQL & vbCrLf & " where acptdd = '" & Trim(strAcptdd) & "'"
            SQL = SQL & vbCrLf & "   and acptno = '" & Trim(strAcptno) & "'"
            SQL = SQL & vbCrLf & "   and testcd_cd = '" & Trim(sExamCode) & "'"
            SQL = SQL & vbCrLf & "   and cnfmflag <> '1' "
'
                  'and testcd_cd = '검사코드'
            res = SendQuery(gServer, SQL)
            If res = -1 Then
                Save_Raw_Data "[QueryErr]" & SQL
                
                'Exit Function
            End If
            
            '/결과 다 넣고 난 뒤에 프로시져임
            'shinrslt..SP_EDIT_RSLT '일자' , '번호','검사코드','결과' , '시간'

        
            Set CmEx = New ADODB.Command
            
            With CmEx
                .ActiveConnection = cn_Ser
                .CommandType = adCmdText
    
                .CommandText = "shinrslt..SP_EDIT_RSLT " & _
                               "'" & Trim(strAcptdd) & "' ," & _
                               "'" & Trim(strAcptno) & "' ," & _
                               "'" & Trim(sExamCode) & "' ," & _
                               "'" & Trim(sTransRes) & "' ," & _
                               "'" & Trim(sTransDate) & Trim(sTransTime) & "'"
                .Execute
            End With
            
            Set CmEx = Nothing
            
            strINT_SNO = Val(strINT_SNO) + 1
            
            Call INSERT_RESULT(lsID, strWKSEQ, strINT_SNO, sTransRes, sExamCode, "")
            
            
        End If
        DoSleep 1
    Next i
    
    SQL = "update pat_res " & vbCrLf & _
          " set sendflag = '2' " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(argSPR, argSpcRow, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
                    
    Insert_Data_SHINWON = 1
    
    Exit Function
    
errhadle:
    Save_Raw_Data "[ERROR] - " & Err.no & "   " & Err.Description
    Save_Raw_Data "[ERROR] - " & SQL
    
    DisConnect_Server
    Connect_Server
    

End Function

Function INSERT_RESULT(argBarcode As String, argWKSEQ As String, argINT_SNO As String, _
                       argResult As String, argTESTCD_CD As String, Optional argposno As String = "")



If Len(argBarcode) <> 15 Then Exit Function

'SELECT A.RGSTNO, A.NM, A.CUSTCD_CD, A.CHARTNO , A.SEX, A.AGE, A.SMOK, B.RVAL
'  FROM shinrslt..acptinfo A, shinrslt..testrslt B
'Where a.ACPTDD = b.ACPTDD
'AND A.ACPTNO = B.ACPTNO
'AND A.ACPTDD = '20121222'
'AND B.ACPTNO = '6140035'

Dim RGSTNO As String
Dim NM As String
Dim CUSTCD_CD As String
Dim CHARTNO As String
Dim Sex As String
Dim Age As String
Dim SMOK As String
Dim RVAL As String
Dim JUDG As String
Dim PANICFLAG As String
Dim DELTAFLAG As String

On Error GoTo ErrHandler

    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.RGSTNO, A.NM, A.CUSTCD_CD, A.CHARTNO , A.SEX, A.AGE, A.SMOK, B.RVAL, B.JUDG, B.PANICFLAG, B.DELTAFLAG "
    SQL = SQL & vbCrLf & "  FROM shinrslt..acptinfo A, shinrslt..testrslt B"
    SQL = SQL & vbCrLf & " Where a.ACPTDD = b.ACPTDD"
    SQL = SQL & vbCrLf & "   AND A.ACPTNO = B.ACPTNO"
    SQL = SQL & vbCrLf & "   AND A.ACPTDD = '" & "20" & Mid(argBarcode, 1, 6) & "'"
    SQL = SQL & vbCrLf & "   AND B.ACPTNO = '" & Mid(argBarcode, 7, 7) & "'"
    SQL = SQL & vbCrLf & "   AND B.TESTCD_CD = '" & argTESTCD_CD & "'"
    SQL = SQL & vbCrLf & " GROUP BY A.RGSTNO, A.NM, A.CUSTCD_CD, A.CHARTNO , A.SEX, A.AGE, A.SMOK, B.RVAL, B.JUDG, B.PANICFLAG, B.DELTAFLAG "
    res = db_select_Col(gServer, SQL)

    RGSTNO = gReadBuf(0)
    NM = gReadBuf(1)
    CUSTCD_CD = gReadBuf(2)
    CHARTNO = gReadBuf(3)
    Sex = gReadBuf(4)
    Age = gReadBuf(5)
    SMOK = gReadBuf(6)
    RVAL = gReadBuf(7)
    JUDG = gReadBuf(8)
    PANICFLAG = gReadBuf(9)
    DELTAFLAG = gReadBuf(10)


    SQL = ""
    SQL = SQL & vbCrLf & "INSERT INTO SHINRSLT..INTRSLT(INT_MACOD, WKDATE, WKSEQ, WKJANG, WKRACK, WKPOS, "
    SQL = SQL & vbCrLf & "                              ACPTDD, ACPTNO, RGSTNO ,NM, CUSTCD_CD, CHARTNO, "
    SQL = SQL & vbCrLf & "                              INT_SNO, TESTCD_CD, RSLT, "
    SQL = SQL & vbCrLf & "                              JUDG, PANICFLAG, DELTAFLAG, RVAL, WKBARCODE, REGSTATE, SEX, AGE, SMOK) "
    SQL = SQL & vbCrLf & "                       values('cobas 8000'"
    SQL = SQL & vbCrLf & "                             ,'" & Format(Date, "yyyymmdd") & "' "
    SQL = SQL & vbCrLf & "                             ,'" & argWKSEQ & "' "
    SQL = SQL & vbCrLf & "                             ,'" & argBarcode & "' "
    SQL = SQL & vbCrLf & "                             ,'" & "' "
    SQL = SQL & vbCrLf & "                             ,'" & argposno & "' "
    
    SQL = SQL & vbCrLf & "                             ,'" & "20" & Mid(argBarcode, 1, 6) & "' "
    SQL = SQL & vbCrLf & "                             ,'" & Mid(argBarcode, 7, 7) & "' "
    SQL = SQL & vbCrLf & "                             ,'" & RGSTNO & "' "
    SQL = SQL & vbCrLf & "                             ,'" & NM & "' "
    SQL = SQL & vbCrLf & "                             ,'" & CUSTCD_CD & "' "
    SQL = SQL & vbCrLf & "                             ,'" & CHARTNO & "' "
    
    SQL = SQL & vbCrLf & "                             ,'" & argINT_SNO & "' "
    SQL = SQL & vbCrLf & "                             ,'" & argTESTCD_CD & "' "
    SQL = SQL & vbCrLf & "                             ,'" & argResult & "' "
    
    SQL = SQL & vbCrLf & "                             ,'" & JUDG & "' "
    SQL = SQL & vbCrLf & "                             ,'" & PANICFLAG & "' "
    SQL = SQL & vbCrLf & "                             ,'" & DELTAFLAG & "' "
    
    SQL = SQL & vbCrLf & "                             ,'" & RVAL & "' "
    SQL = SQL & vbCrLf & "                             ,'" & argBarcode & "' "
    SQL = SQL & vbCrLf & "                             ,'' "
    SQL = SQL & vbCrLf & "                             ,'" & Sex & "' "
    SQL = SQL & vbCrLf & "                             ,'" & Age & "' "
    SQL = SQL & vbCrLf & "                             ,'N')"
    res = SendQuery(gServer, SQL)
    
    Exit Function

ErrHandler:
    
    Save_Raw_Data "ERR - " & argBarcode & " : " & argTESTCD_CD & vbCrLf & SQL


End Function

Function Insert_Data(argSpcRow As Integer, argSpread As vaSpread, Optional argEXAMDATE As String = "", Optional argFSendFlag As String = "", Optional argAoutoYN As String = "N") As Integer
    Dim iRow        As Integer
    Dim i           As Integer
    Dim X           As Integer
    Dim j           As Integer
    Dim lsID        As String
    Dim lsPosNo     As String
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
    Dim sExamCodeList As String
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
    Dim lsUnit As String
    Dim strRes As String
    Dim lsOrderCode As String
    Dim lsSubCode As String
    Dim strLabTestIDNo As String
    
    Dim strLabTestState As String
    Dim intLabTestRow As Integer
    Dim intResFalg As Integer
    
    Dim strBeforResult  As String
    
    '최종보고때문에 만듬
    Dim intFSendCode As Integer
    Dim strFSendCode As String
    Dim arrFSendCode(0 To 30) As String
    
    Dim strResCode As String
    
    
    Dim strRsltNum As String
    
    Dim strSendYN As String
    Dim strDept As String
    Insert_Data = -1
    
    intFSendCode = 0
    
    sResFlag = False
    lsID = ""
    lsID = Trim(GetText(argSpread, argSpcRow, colBarCode))
    lsPosNo = Trim(GetText(argSpread, argSpcRow, colRackPos))
    
    Dim intTransCnt As Integer
    intTransCnt = 0
    
    Dim sRes As String
    Dim strSpcDate As String
    Dim strSpcNo As String
    Dim strSpcSeq As String
    SQL = "SELECT SPCDATE, SPCNO, SPCSEQ "
    SQL = SQL & vbCrLf & "  From SLACPTMT "
    SQL = SQL & vbCrLf & " WHERE SPCDATE = TO_DATE('" & Mid(lsID, 1, 6) & "', 'YYMMDD') "
    SQL = SQL & vbCrLf & "   AND SPCNO = '" & Mid(lsID, 7, 5) & "'"
    SQL = SQL & vbCrLf & "   AND SPCSEQ = '" & Mid(lsID, 12, 1) & "'"
    res = db_select_Col(gServer, SQL)
    
    If res = 0 Then
        Exit Function
    End If
    
    
    strSpcDate = Trim(gReadBuf(0))
    strSpcNo = Trim(gReadBuf(1))
    strSpcSeq = Trim(gReadBuf(2))
    'EXAMEUID에 처방과 정보를 넣어둠.
    SQL = " Select EXAMUID"
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & lsID & "'"
    If argEXAMDATE <> "" Then
        SQL = SQL & vbCrLf & "AND EXAMDATE = '" & argEXAMDATE & "'"
    End If
    'SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & argEXAMDATE & "'"
    SQL = SQL & vbCrLf & "   AND EXAMUID <> '' "
    res = db_select_Col(gLocal, SQL)
    
    strDept = Trim(gReadBuf(0))
    
    '160408075850
    
    
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasResTemp
    
    SQL = " Select A.equipcode, A.examcode, A.resvalue, A.result, A.refflag, B.DELTAVALUE " & vbCrLf & _
          " From pat_res A, EQUIPEXAM B  " & vbCrLf & _
          " Where A.barcode = '" & lsID & "' and A.resvalue <> '' "
    SQL = SQL & vbCrLf & "AND A.EQUIPCODE = B.EQUIPCODE"
    
    If argEXAMDATE <> "" Then
        SQL = SQL & vbCrLf & "AND A.EXAMDATE = '" & argEXAMDATE & "'"
    End If
    SQL = SQL & vbCrLf & " GROUP BY A.equipcode, A.examcode, A.resvalue, A.result, A.refflag, B.DELTAVALUE"
    SQL = SQL & vbCrLf & "ORDER BY B.DELTAVALUE DESC"
    
    'SQL = SQL & vbCrLf & "AND POSNO = '" & lsPosNo & "'"
    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    '자동전송일때 전송여부
    strSendYN = "Y"
    
    Save_Raw_Data "[서버전송 시작 " & Format(Now, "hh:mm:ss") & "]"
    
    For j = 1 To vasResTemp.DataRowCnt
        strResCode = ""
        lsOrderCode = ""
        lsSubCode = ""
        strBeforResult = ""
        strLabTestIDNo = ""
        strLabTestState = ""
        sEquipCode = Trim(GetText(vasResTemp, j, 1))
        sExamCode = Trim(GetText(vasResTemp, j, 2))
        sResValue = Trim(GetText(vasResTemp, j, 3))
        sResult = Trim(GetText(vasResTemp, j, 4))
        sRefFlag = Trim(GetText(vasResTemp, j, 5))
        
        
        '아래의 조건일경우 MICRO결과를 자동전송하지 않는다.
        If argAoutoYN = "Y" Then
            If sEquipCode = "ERY" Then
                If UCase(sResValue) = "TRACE" Then
                    strSendYN = "N"
                ElseIf UCase(sResValue) = "1+" Then
                    strSendYN = "N"
                ElseIf UCase(sResValue) = "2+" Then
                    strSendYN = "N"
                ElseIf UCase(sResValue) = "3+" Then
                    strSendYN = "N"
                End If
                
            ElseIf sEquipCode = "LEU" Then
                If UCase(sResValue) = "1+" Then
                    strSendYN = "N"
                ElseIf UCase(sResValue) = "2+" Then
                    strSendYN = "N"
                ElseIf UCase(sResValue) = "3+" Then
                    strSendYN = "N"
                End If
                
            ElseIf sEquipCode = "PRO" Then
                If UCase(sResValue) = "1+" Then
                    strSendYN = "N"
                ElseIf UCase(sResValue) = "2+" Then
                    strSendYN = "N"
                ElseIf UCase(sResValue) = "3+" Then
                    strSendYN = "N"
                End If
                
            ElseIf sEquipCode = "NIT" Then
                If UCase(sResValue) = "POS" Then
                    strSendYN = "N"
                End If
            End If
        End If
        
        
        If sResValue <> "" And sResValue <> "-" Then
            
            
            
            '결과코드 설정.
            '하드코딩하기로함.
            If UCase(sResValue) = "NEGATIVE" Then
                strResCode = "N"
            ElseIf UCase(sResValue) = "NEG" Then
                strResCode = "N"
                
            ElseIf UCase(sResValue) = "POS" Then
                strResCode = "P"
            ElseIf UCase(sResValue) = "POSITIVE" Then
                strResCode = "P"
            
            ElseIf UCase(sResValue) = "1+" Then
                strResCode = "1P"
            ElseIf UCase(sResValue) = "2+" Then
                strResCode = "2P"
            ElseIf UCase(sResValue) = "3+" Then
                strResCode = "3P"
            ElseIf UCase(sResValue) = "4+" Then
                strResCode = "4P"
            ElseIf UCase(sResValue) = "5+" Then
                strResCode = "5P"
                
            ElseIf UCase(sResValue) = "TRACE" Then
                strResCode = "TR"
            
            ElseIf UCase(sResValue) = "TRACE" Then
                strResCode = "TR"
            ElseIf UCase(sResValue) = "TR" Then
                strResCode = "TR"
            
            ElseIf UCase(sResValue) = "P.YEL" Then
                strResCode = "PR"
            ElseIf UCase(sResValue) = "YELLOW" Then
                strResCode = "YE"
            ElseIf UCase(sResValue) = "AMBER" Then
                strResCode = "AM"
            ElseIf UCase(sResValue) = "BROWN" Then
                strResCode = "BR"
            ElseIf UCase(sResValue) = "ORANGE" Then
                strResCode = "OR"
            ElseIf UCase(sResValue) = "RED" Then
                strResCode = "RE"
            ElseIf UCase(sResValue) = "GREEN" Then
                strResCode = "GR"
            ElseIf UCase(sResValue) = "OTHER" Then
                strResCode = "OT"
            ElseIf UCase(sResValue) = "CLEAR" Then
                strResCode = "T1"
            ElseIf UCase(sResValue) = "TURBID" Then
                strResCode = "T4"
            ElseIf UCase(sResValue) = "MUCOUS" Then
                strResCode = "T5"
            ElseIf UCase(sResValue) = "SL.CLOUDY" Then
                strResCode = "T6"
            ElseIf UCase(sResValue) = "VERY.CLOUDY" Then
                strResCode = "T7"
            ElseIf UCase(sResValue) = "BLOODY" Then
                strResCode = "T8"
            ElseIf UCase(sResValue) = "≥50" Then
                strResCode = "M"
            Else
                strResCode = ""
            End If
            
            'SG의 경우 장비값과 결과값을 다르게 넣음
            'SG의 경우 예외적으로 IF 하나 더씀
            strRsltNum = ""
            If sEquipCode = "SG" Then
                If IsNumeric(sResValue) = True Then
                    If CCur(sResValue) >= "1.050" Then
                        strResCode = "H"
                        strRsltNum = "≥1.050"
                    End If
                End If
            End If
            
            
            '결과 형태 확인.
            SQL = ""
            SQL = SQL & vbCrLf & "SELECT deltavalue"
            SQL = SQL & vbCrLf & "  FROM EQUIPEXAM"
            SQL = SQL & vbCrLf & " WHERE EQUIPCODE = '" & sEquipCode & "'"
            SQL = SQL & vbCrLf & "GROUP BY deltavalue"
            SQL = SQL & vbCrLf & ""
            SQL = SQL & vbCrLf & ""
            res = db_select_Col(gLocal, SQL)
            
            If gReadBuf(0) = "1" Then
                If sResValue = "Neg" Then strResCode = ""
                If sResValue = "Pos" Then strResCode = ""
                If strRsltNum <> "" Then
                    sRes = spUpdateResult(strSpcDate, strSpcNo, strSpcSeq, sEquipCode, strResCode, sResValue, strRsltNum, "", gEquip, _
                                              "POCT", Winsock2.LocalIP, "", "", "", "", "")
                
                Else
                    sRes = spUpdateResult(strSpcDate, strSpcNo, strSpcSeq, sEquipCode, strResCode, sResValue, sResValue, "", gEquip, _
                                              "POCT", Winsock2.LocalIP, "", "", "", "", "")
                End If
                intTransCnt = intTransCnt + 1
                
                SQL = "update pat_res " & vbCrLf & _
                      " set sendflag = '2' " & vbCrLf & _
                      " Where equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(argSpread, argSpcRow, colBarCode)) & "' " & vbCrLf & _
                      " AND EQUIPCODE = '" & Trim(sEquipCode) & "'"
                res = SendQuery(gLocal, SQL)
                If res < 0 Then
                    'cn_Ser.RollbackTrans
                    Save_Raw_Data SQL
                    Exit Function
                End If
            Else
                
                If argAoutoYN = "Y" And strSendYN = "Y" Then
                    '자동전송이고, 전송여부가 Y일때만 자동전송한다.
                    '경우에 따라 전송하는 조건이 달라짐.
                    If argFSendFlag = "Y" Then
                        '자동이고 경우에 따라 침사결과 정상이고 검경결과가 이상일때 혹은 반대의 상황이 될때
                        '검경결과를 자동전송 하지 않음.
                        'sRes = spUpdateResult(strSpcDate, strSpcNo, strSpcSeq, sEquipCode, strResCode, sResValue, sResValue, "", gEquip, _
                        '                       "POCT", Winsock1.LocalIP, "", "", "", "", "")
                    ElseIf argFSendFlag = "YY" Then
                        intTransCnt = intTransCnt + 1
                        If sEquipCode = "WBC" Then
                            sRes = spUpdateResult(strSpcDate, strSpcNo, strSpcSeq, sEquipCode, "0", "0", "0", "", gEquip, _
                                                  "POCT", Winsock1.LocalIP, "", "", "", "", "")
                        ElseIf sEquipCode = "RBC" Then
                            sRes = spUpdateResult(strSpcDate, strSpcNo, strSpcSeq, sEquipCode, "0", "0", "0", "", gEquip, _
                                                  "POCT", Winsock1.LocalIP, "", "", "", "", "")
                        Else
                            sRes = spUpdateResult(strSpcDate, strSpcNo, strSpcSeq, sEquipCode, " ", " ", " ", "", gEquip, _
                                                  "POCT", Winsock1.LocalIP, "", "", "", "", "")
                        End If
                        SQL = "update pat_res " & vbCrLf & _
                              " set sendflag = '2' " & vbCrLf & _
                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
                              " And barcode = '" & Trim(GetText(argSpread, argSpcRow, colBarCode)) & "' " & vbCrLf & _
                              " AND EQUIPCODE = '" & Trim(sEquipCode) & "'"
                        res = SendQuery(gLocal, SQL)
                        If res < 0 Then
                            'cn_Ser.RollbackTrans
                            Save_Raw_Data SQL
                            Exit Function
                        End If
                    End If
                    
                ElseIf argAoutoYN = "N" Then
                    intTransCnt = intTransCnt + 1
                    '자동전송이 아니고 전송여부가 Y일때는 그냥 전부 전송한다.
                    sRes = spUpdateResult(strSpcDate, strSpcNo, strSpcSeq, sEquipCode, strResCode, sResValue, sResValue, "", gEquip, _
                                               lblSendID.Caption, Winsock1.LocalIP, "", "", "", "", "")
                                               
                    SQL = "update pat_res " & vbCrLf & _
                          " set sendflag = '2' " & vbCrLf & _
                          " Where equipno = '" & gEquip & "' " & vbCrLf & _
                          " And barcode = '" & Trim(GetText(argSpread, argSpcRow, colBarCode)) & "' " & vbCrLf & _
                          " AND EQUIPCODE = '" & Trim(sEquipCode) & "'"
                    res = SendQuery(gLocal, SQL)
                    If res < 0 Then
                        'cn_Ser.RollbackTrans
                        Save_Raw_Data SQL
                        Exit Function
                    End If
                End If
            End If
            'Save_Raw_Data sRes
            
            
            
        End If
        
        
    Next j
    
    
    
'''    SQL = "update pat_res " & vbCrLf & _
'''          " set sendflag = '2' " & vbCrLf & _
'''          " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''          " And barcode = '" & Trim(GetText(argSpread, argSpcRow, colBarCode)) & "' " & vbCrLf & _
'''          " And sendflag NOT IN ( '5','6') "
'''    res = SendQuery(gLocal, SQL)
    
    Save_Raw_Data "[서버전송 완료 " & Format(Now, "hh:mm:ss") & "]"

    If Trim(GetText(argSpread, argSpcRow, colState)) = "미접수" Then
        SQL = "update pat_res " & vbCrLf & _
            " set POSNO = '" & Trim(GetText(argSpread, argSpcRow, colRackPos)) & "' " & vbCrLf & _
            " Where equipno = '" & gEquip & "' " & vbCrLf & _
            " And barcode = '" & Trim(GetText(argSpread, argSpcRow, colBarCode)) & "' " & vbCrLf & _
            " And sendflag NOT IN ( '5','6') "
        res = SendQuery(gLocal, SQL)
    ElseIf Trim(GetText(argSpread, argSpcRow, colState)) = "오더없음" Then
        SQL = "update pat_res " & vbCrLf & _
            " set POSNO = '" & Trim(GetText(argSpread, argSpcRow, colRackPos)) & "' " & vbCrLf & _
            " Where equipno = '" & gEquip & "' " & vbCrLf & _
            " And barcode = '" & Trim(GetText(argSpread, argSpcRow, colBarCode)) & "' " & vbCrLf & _
            " And sendflag NOT IN ( '5','6') "
        res = SendQuery(gLocal, SQL)
    End If
    If intTransCnt > 0 Then
        Insert_Data = 1
    End If
End Function

Function Insert_Data_ADD(argSpcRow As Integer, argSpread As vaSpread) As Integer
    Dim iRow        As Integer
    Dim i           As Integer
    Dim X           As Integer
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
    Dim sExamCodeList As String
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
    Dim lsUnit As String
    Dim strRes As String
    Dim lsOrderCode As String
    Dim lsSubCode As String
    Dim intResFalg As Integer
    
    
    
    Dim strBeforResult  As String
    
    Insert_Data_ADD = -1
    
    sResFlag = False
    lsID = ""
    lsID = Trim(GetText(argSpread, argSpcRow, colBarCode))
    
    
    lsReceDate = "20" & Mid(lsID, 1, 2) & "/" & Mid(lsID, 3, 2) & "/" & Mid(lsID, 5, 2)
    lsReceNo = Mid(lsID, 7, 6)
    If IsNumeric(lsReceNo) = False Then
        Exit Function
    End If

    lsReceNo = CStr(CCur(lsReceNo))
    
    ClearSpread vasTransTemp
    
    SQL = "Interface_GetPatientResult '" & gWorkList & "', '" & lsReceDate & "', " & lsReceNo & ""
    res = db_select_VasS(gServer, SQL, vasTransTemp)
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasResTemp
    
    SQL = " Select equipcode, examcode, resvalue, result, refflag " & vbCrLf & _
          " From pat_res  " & vbCrLf & _
          " Where barcode = '" & lsID & "' and resvalue <> '' AND DeltaFlag = '*' "

    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
        
    For j = 1 To vasResTemp.DataRowCnt
        lsOrderCode = ""
        lsSubCode = ""
        strBeforResult = ""
        sEquipCode = ""
        
        sEquipCode = Trim(GetText(vasResTemp, j, 1))
        sExamCode = Trim(GetText(vasResTemp, j, 2))
        sResValue = Trim(GetText(vasResTemp, j, 3))
        sResult = Trim(GetText(vasResTemp, j, 4))
        sRefFlag = Trim(GetText(vasResTemp, j, 5))
        
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT EXAMCODE"
        SQL = SQL & vbCrLf & "  FROM EQUIPEXAM"
        SQL = SQL & vbCrLf & " WHERE EQUIPCODE = '" & sEquipCode & "'"
        res = db_select_Row(gLocal, SQL)
        sExamCodeList = "''"
        For X = 0 To res - 1
            If sExamCodeList = "''" Then
                sExamCodeList = "'" & gReadBuf(X) & "'"
            Else
                sExamCodeList = sExamCodeList & ",'" & gReadBuf(X) & "'"
            End If
        Next X
        
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT TESTSUBCODE, TESTCODE, ORDERCODE"
        SQL = SQL & vbCrLf & "  FROM LabRegResult "
        SQL = SQL & vbCrLf & " WHERE LabRegDate = '" & lsReceDate & "'"
        SQL = SQL & vbCrLf & "   AND LabRegNo = '" & Trim(lsReceNo) & "'"
        SQL = SQL & vbCrLf & "   AND TESTSUBCODE IN (" & sExamCodeList & ")"
        res = db_select_Col(gServer, SQL)
        
        lsSubCode = gReadBuf(0)
        sExamCode = gReadBuf(1)
        lsOrderCode = gReadBuf(2)
'        For X = 1 To vasTransTemp.DataRowCnt
'            If Trim(GetText(vasTransTemp, X, 11)) = sExamCode Then
'                lsOrderCode = Trim(GetText(vasTransTemp, X, 9))
'                sExamCode = Trim(GetText(vasTransTemp, X, 10))
'                lsSubCode = Trim(GetText(vasTransTemp, X, 11))
'                '현재 서버에 있는 결과값
'                strBeforResult = Trim(GetText(vasTransTemp, X, 13))
'            End If
'
'        Next
        
        If lsOrderCode = "" Then lsOrderCode = sExamCode
        If lsSubCode = "" Then lsSubCode = sExamCode
        
        SQL = "select resgubun, '' from equipexam where equipcode = '" & Trim(GetText(vasResTemp, j, 1)) & "'"
        res = db_select_Col(gLocal, SQL)
        
        sResGubun = Trim(gReadBuf(0))
        lsUnit = Trim(gReadBuf(1))
        
'''        If sResGubun = "1" Then '문자
'''
'''            sTransRes = sResult & "(" & sResValue & ")"
'''        Else
'''            sTransRes = sResValue
'''            sResult = ""
'''        End If
        
        'If strBeforResult = "" Then
            SQL = "Interface_SetPatientResult '" & lsReceDate & "', " & lsReceNo & ", " & _
                  "'" & lsOrderCode & "', '" & sExamCode & "', '" & lsSubCode & "', '" & sResValue & "', " & _
                  "'" & "" & "', '" & sRefFlag & "', 0, 0, 0, '" & gEquip & "', " & intResFalg & " "
            res = SendQuery(gServer, SQL)
        'End If
        
'''        SQL = "UPDATE LC04_LABGE.dbo.LabRegResult"
'''        SQL = SQL & vbCrLf & "SET TestResultAbn = '" & sRefFlag & "'"
'''        SQL = SQL & vbCrLf & " WHERE LabRegDate = '" & lsReceDate & "'"
'''        SQL = SQL & vbCrLf & "   AND LabRegNo = '" & lsReceNo & "'"
'''        SQL = SQL & vbCrLf & "   AND OrderCode = '" & lsOrderCode & "'"
'''        SQL = SQL & vbCrLf & "   AND TestCode = '" & sExamCode & "'"
'''        SQL = SQL & vbCrLf & "   AND TestSubCode = '" & lsSubCode & "'"
'''        res = SendQuery(gServer, SQL)
        
            
        
'''        If Mid(lsID, 1, 2) = "99" Then
'''            SQL = " UPDATE  SLAXEQRDT  "
'''            SQL = SQL & vbCrLf & " SET EXAM_RSLT = '" & sResValue & "' , USE_YN = 'Y' "
'''            SQL = SQL & vbCrLf & "  where  BAR_NO = '" & lsID & "'  AND EXAM_CD = '" & sExamCode & "' "
'''            SQL = SQL & vbCrLf & "    and (trim(EXAM_RSLT) = '' or exam_rslt is null)   "
'''            res = SendQuery(gServer, SQL)
'''
'''        Else
'''
'''            strRes = spUpdateResult(lsID, sExamCode, sTransRes, "", "D", gEquip, "", sRefFlag)
'''        End If
        
        
    Next
    
'''    ClearSpread vasRemark
'''
'''    SQL = "select sremark from serum_index where barcode = '" & lsID & "' and sremark <> '' "
'''    res = db_select_Vas(gLocal, SQL, vasRemark)
'''
'''    For j = 1 To vasRemark.DataRowCnt
'''        If Trim(GetText(vasRemark, j, 1)) <> "" Then
'''            strRes = spUpdateRemark(lsID, Trim(GetText(vasRemark, j, 1)))
'''        End If
'''    Next
'''
 
    SQL = "update pat_res " & vbCrLf & _
          " set sendflag = '2' " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(argSpread, argSpcRow, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
                    
    Insert_Data_ADD = 1
    
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
        
    
    Insert_Data_1 = -1
    
    lsID = ""
    lsID = Trim(GetText(vasList, argSpcRow, colBarCode))
    
    sTransDate = Format(GetDateFull, "yyyymmdd")
    sTransTime = Format(GetDateFull, "hhmmss")
    
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasTemp
    
    SQL = " Select a.equipcode, a.examcode, a.resvalue, a.result, b.resgubun " & vbCrLf & _
          " From pat_res a, equipexam b " & vbCrLf & _
          " Where a.equipno = b.equipno " & vbCrLf & _
          " And a.examcode = b.examcode " & vbCrLf & _
          " And a.equipcode = b.equipcode " & vbCrLf & _
          " And a.equipno = '" & gEquip & "' " & vbCrLf & _
          " And a.barcode = '" & lsID & "' "
          
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    sCnt = ""
    cn_Ser.BeginTrans
    '서버로 결과값 저장하기
    For i = 1 To vasTemp.DataRowCnt
        
            
        sExamCode = Trim(GetText(vasTemp, i, 2))
        sResValue = Trim(GetText(vasTemp, i, 3))
        sResult = Trim(GetText(vasTemp, i, 4))
        sResGubun = Trim(GetText(vasTemp, i, 5))
        
        If sResGubun = "1" Then '문자
            sTransRes = sResValue & "(" & sResult & ")"
            
        Else
            sTransRes = sResValue
            sResult = ""
        End If
        
        
        If sExamCode <> "" And sResValue <> "" Then

            SQL = "SELECT A.SPCM_NO, AA.RSLT_SQNO, A.RCPN_SQNO " & vbCrLf & _
                  "FROM MS.MSLRCPT A " & vbCrLf & _
                  "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
                  "WHERE A.SPCM_NO = '" & lsID & "' " & vbCrLf & _
                  "  AND AA.EXMN_CD = '" & sExamCode & "'"
            
            res = db_select_Col(gServer, SQL)
            If res = -1 Then
                Save_Raw_Data "[QueryErr]" & SQL
                Exit Function
                
            End If
            
            sRsltSqno = Trim(gReadBuf(1))
            sRcpnSqno = Trim(gReadBuf(2))
            '/아래 조건이 어긋나면 전송 취소
            If Trim(sRsltSqno) <> "" And Trim(sRcpnSqno) <> "" Then
            
                SQL = "select eqpm_rslt_valu from mslintrslt " & vbCrLf & _
                          " where rslt_sqno = '" & sRsltSqno & "' "
                    res = db_select_Col(gServer, SQL)
                If res = -1 Then
                    Save_Raw_Data "[QueryErr]" & SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                End If
                    
                 If res > 0 Then
                                    SQL = "UPDATE MSLINTRSLT"
                     SQL = SQL & vbCrLf & "   SET EQPM_RSLT_VALU = '" & sResValue & "'"
                     SQL = SQL & vbCrLf & "      ,INIT_EQPM_RSLT_VALU = '" & sTransRes & "'"
                     SQL = SQL & vbCrLf & "      ,RSLT_PRGR_STAT_CD = '07'"
                     SQL = SQL & vbCrLf & "      ,LAST_UPDT_USID = '" & gExamUID & "'"
                     SQL = SQL & vbCrLf & "      ,LAST_UDDT = SYSDATE "
                     SQL = SQL & vbCrLf & " WHERE RSLT_SQNO = '" & sRsltSqno & "'"
                     SQL = SQL & vbCrLf & "   AND EQPM_RCPN_SQNO = '" & sRcpnSqno & "'"
                     SQL = SQL & vbCrLf & "   AND RSLT_PRGR_STAT_CD < '11'"
    
                     res = SendQuery(gServer, SQL)
                     If res = -1 Then
                         Save_Raw_Data "[QueryErr]" & SQL
                         cn_Ser.RollbackTrans
                         Exit Function
                     End If
                 ElseIf res = 0 Then
                     SQL = "insert into mslintrslt (rslt_sqno, rslt_trms_date, rslt_trms_time, eqpm_cd, eqpm_rslt_valu, " & vbCrLf & _
                           "eqpm_rslt_dvcd, err_valu, init_eqpm_rslt_valu, updt_eqpm_rslt_valu, eqpm_rslt_rmrk, " & vbCrLf & _
                           "eqpm_rcpn_sqno, rslt_prgr_stat_cd, frst_rgst_usid, frst_rgdt, last_updt_usid, last_uddt) " & vbCrLf & _
                           "values( " & vbCrLf & _
                           "'" & sRsltSqno & "','" & sResValue & "','" & sTransTime & "', " & vbCrLf & _
                           "'" & gEquip & "','" & sTransRes & "', " & vbCrLf & _
                           "'','','" & sTransRes & "', " & vbCrLf & _
                           "'','', " & vbCrLf & _
                           "'" & sRcpnSqno & "','07', '" & gExamUID & "', " & vbCrLf & _
                           "SYSDATE,'" & gExamUID & "',SYSDATE " & vbCrLf & _
                           ") "
                     res = SendQuery(gServer, SQL)
                     If res = -1 Then
                         Save_Raw_Data "[QueryErr]" & SQL
                         cn_Ser.RollbackTrans
                         Exit Function
                         
                     End If
                End If
                
                SQL = "UPDATE MS.MSLGNRLRSLT " & vbCrLf & _
                      "SET    RSLT_PRGR_STAT_CD = '07',  --결과저장(예비결과)  " & vbCrLf & _
                      "       NMVL_RSLT_VALU = '" & sResValue & "',  " & vbCrLf & _
                      "       TXT_RSLT_VALU = '" & sTransRes & "', " & vbCrLf & _
                      "       NRML_DVCD = '', " & vbCrLf & _
                      "       DELT_YN = '', " & vbCrLf & _
                      "       PANC_YN = '', " & vbCrLf & _
                      "       ALRT_YN = '', " & vbCrLf & _
                      "       EXMN_RSLT_STOR_DATE = TO_CHAR(SYSDATE, 'YYYYMMDD'), " & vbCrLf & _
                      "       EXMN_RSLT_STOR_TIME = TO_CHAR(SYSDATE, 'HH24MISS'), " & vbCrLf & _
                      "       EXMN_RSLT_STOR_PRSN_ID = '" & gExamUID & "', " & vbCrLf & _
                      "       LAST_UPDT_USID = '" & gExamUID & "', " & vbCrLf & _
                      "       LAST_UDDT = SYSTIMESTAMP, EXMN_EQPM_CD = '" & gEquip & "'  " & vbCrLf & _
                      " WHERE RSLT_SQNO = '" & sRsltSqno & "' AND RSLT_PRGR_STAT_CD <> '11' "
                res = SendQuery(gServer, SQL)
                
                If res = -1 Then
                    Save_Raw_Data "[QueryErr]" & SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                    
                End If
                
                SQL = "UPDATE MS.MSLRCPT " & vbCrLf & _
                      " SET   exmn_prgr_stat_cd = '07', " & vbCrLf & _
                      "        last_updt_usid = '" & gExamUID & "', " & vbCrLf & _
                      "        last_uddt = SYSTIMESTAMP " & vbCrLf & _
                      "  WHERE RCPN_SQNO = '" & sRcpnSqno & "' "
                res = SendQuery(gServer, SQL)
                
                If res = -1 Then
                    Save_Raw_Data "[QueryErr]" & SQL
                    cn_Ser.RollbackTrans
                    Exit Function
                    
                End If
            End If
        End If
        DoSleep 50
    Next i
    cn_Ser.CommitTrans
    
    SQL = "update pat_res " & vbCrLf & _
          " set sendflag = '2' " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasList, argSpcRow, colBarCode)) & "' "
    res = SendQuery(gLocal, SQL)
                    
    Insert_Data_1 = 1
    
End Function

Private Sub cmdAbnormal_Click()
    frmExamSearch.Show
End Sub

Private Sub cmdAddDataCall_Click()
    Dim i As Integer
    ClearSpread vasList
    
    SQL = "select '', barcode, posno, pid, pname, PSEX, count(result), sendflag from pat_res " & vbCrLf & _
          " where examdate Between '" & Format(dtpAddExamDate1, "yyyymmdd") & "' AND '" & Format(dtpAddExamDate2, "yyyymmdd") & "'"
    SQL = SQL & " and deltaflag = '" & "*" & "' "
    
    SQL = SQL & vbCrLf & " group by  barcode, posno, pid, pname, PSEX, sendflag "
    res = db_select_Vas(gLocal, SQL, vasList)

    
    vasList.MaxRows = vasList.DataRowCnt
    For i = 1 To vasList.DataRowCnt
        If GetText(vasList, i, colState) = "1" Then
            SetText vasList, "Result", i, colState
            
        ElseIf GetText(vasList, i, colState) = "2" Then
            SetText vasList, "Trans", i, colState
            SetBackColor vasList, i, i, colBarCode, colPSex, 255, 255, 180
            SetBackColor vasList, i, i, colState, colState, 255, 255, 180
        End If
    Next
    
    ClearSpread vasResTemp
End Sub

Private Sub cmdBARERR_Click()
    '18060
    If frmInterface.Width = 18060 Then
        frmInterface.Width = 15315
    Else
        frmInterface.Width = 18060
    End If
End Sub

Private Sub cmdCall_Click()
    Dim i As Long
    Dim varSendFlag
    Dim j As Long
    Dim X As Long
    Dim strResult As String
    
    
    ClearSpread vasList
    
    varSendFlag = cmbTransGubun.ListIndex

    SQL = "select '', FORMAT(examdate + MAX(EXAMTIME), '@@@@-@@-@@ @@:@@:@@'),  posno, MID(TRIM(receno),3), barcode,  pid,  pname, PSEX,  TRIM(EXAMUID), '','',EXAMTYPE from pat_res "
    SQL = SQL & " where examdate BETWEEN '" & Format(dtpExamDate, "yyyymmdd") & "' AND '" & Format(dtpExamDate2, "yyyymmdd") & "'"
          'dtpExamDate2
    If varSendFlag = 1 Or varSendFlag = 2 Then
        SQL = SQL & " and sendflag = '" & varSendFlag & "' "
    ElseIf varSendFlag = 3 Then
        SQL = SQL & " and sendflag = '0' "
    Else
        'SQL = SQL & " and sendflag <> '0' "
    End If
''
''    '신장내과 대상조회
''    If cnkNEMD.Value = 1 Then
''        SQL = SQL & vbCrLf & " ANd TRIM(EXAMUID) LIKE('%NEMD-%')"
''    End If
''
''
''    'Micro 대상조회
''    If chkMicro.Value = 1 Then
''        SQL = SQL & vbCrLf & " ANd EXAMTYPE = 'Y'"
''    End If
    
    If cboMicro.ListIndex = 1 Then
        SQL = SQL & vbCrLf & " ANd EXAMTYPE = '" & cboMicro.Text & "'"
    ElseIf cboMicro.ListIndex = 2 Then
        SQL = SQL & vbCrLf & " ANd EXAMTYPE = '" & cboMicro.Text & "'"
    ElseIf cboMicro.ListIndex = 3 Then
        SQL = SQL & vbCrLf & " ANd (EXAMTYPE IS NULL OR EXAMTYPE = '')"
    End If

    SQL = SQL & vbCrLf & " group by  barcode, posno, pid,MID(TRIM(receno),3), pname, PSEX , TRIM(EXAMUID),EXAMTYPE, examdate"
    res = db_select_Vas(gLocal, SQL, vasList)

    
    vasList.MaxRows = vasList.DataRowCnt
    For i = 1 To vasList.DataRowCnt
        vasList.Row = i
        vasList.Col = -1
        vasList.FontSize = 10
        vasList.RowHeight(i) = 18
        SetVasColor vasList, CInt(i), Trim(GetText(vasList, i, colBarCode))
        If GetText(vasList, i, colState) = "1" Then
            SetText vasList, "Result", i, colState
            
        ElseIf GetText(vasList, i, colState) = "2" Then
            SetText vasList, "Trans", i, colState
            SetBackColor vasList, i, i, colBarCode, colState, 255, 255, 180
        ElseIf GetText(vasList, i, colState) = "5" Then
            SetText vasList, "미접수", i, colState
            SetBackColor vasList, i, i, colBarCode, colState, 50, 50, 255
        ElseIf GetText(vasList, i, colState) = "6" Then
            SetText vasList, "오더없음", i, colState
            SetBackColor vasList, i, i, colBarCode, colState, 50, 50, 255
        End If
    Next
    
    ClearSpread vasResTemp
    
'    SQL = "select barcode, equipcode, resvalue, result from pat_res " & vbCrLf & _
'          " where examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' "
'    If varSendFlag = 1 Or varSendFlag = 2 Then
'        SQL = SQL & " and sendflag = '" & varSendFlag & "' "
'    Else
'        SQL = SQL & " and sendflag <> '0' "
'    End If
'
'    SQL = SQL & vbCrLf & " group by barcode, equipcode, resvalue, result"
'    res = db_select_Vas(gLocal, SQL, vasResTemp)
'
''''    gArr_Exam(i, 1) = i    '순서
''''    gArr_Exam(i, 2) = Trim(GetText(vasTemp, i, 2))    '장비코드
''''    gArr_Exam(i, 3) = Trim(GetText(vasTemp, i, 3))    '검사명
'
'    For i = 1 To vasResTemp.DataRowCnt
'        For j = 1 To vasList.DataRowCnt
'            If Trim(GetText(vasResTemp, i, 1)) = Trim(GetText(vasList, j, colBarCode)) Then
'                For X = 1 To vasList.MaxCols - colRStart
'                    If Trim(GetText(vasResTemp, i, 2)) = Trim(gArr_Exam(X, 2)) Then
'                        If gArr_Exam(X, 4) = "0" Then
'                            strResult = Trim(GetText(vasResTemp, i, 3))
'                        ElseIf gArr_Exam(X, 4) = "1" Then
'                            strResult = Trim(GetText(vasResTemp, i, 4)) & "(" & Trim(GetText(vasResTemp, i, 3)) & ")"
'                        Else
'                            strResult = Trim(GetText(vasResTemp, i, 3))
'                        End If
'
'                        SetText vasList, strResult, j, colRStart + CCur(gArr_Exam(X, 1))
'                        Exit For
'                    End If
'                Next X
'                Exit For
'            End If
'        Next j
'    Next i

End Sub

Private Sub cmdClear_Click()
Dim iNumber As Integer
    
    txtMsg.Text = ""
    
    ClearSpread vasID, 1, 1
    vasID.MaxRows = 0
    ClearSpread vasRes, 1, 1
    vasRes.MaxRows = 0
End Sub


Private Sub cmdLBCLEAR_Click()
    vasBARERR.MaxRows = 0
    cmdBARERR.Caption = "BARCODE ERR"
End Sub

Private Sub cmdListClear_Click()
    Dim iNumber As Integer
    
    txtMsg.Text = ""
    
    ClearSpread vasList, 1, 1
    vasList.MaxRows = 0
    ClearSpread vasLISTRES, 1, 1
    vasLISTRES.MaxRows = 0
End Sub

Private Sub cmdListTrans_Click()
'선택전송
Dim vasIDRow As Integer
Dim VasResRow As Integer
Dim iRow As Integer
Dim liRet As Integer
Dim iNumber As Integer

    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:선택전송") = vbCancel Then
        Exit Sub
    End If
    
    If vasList.DataRowCnt < 1 Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    'db_BeginTran gServer
'''    Connect_Server
    For vasIDRow = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = vasIDRow
        
        If vasList.Value = 1 Then '체크된 열은 저장이 안됨
'        If vasID.Value = "" Then
        
            liRet = -1
            liRet = Insert_Data(vasIDRow, vasList)
            '''liRet = ToServer_Re(vasIDRow)
'''            liRet = Insert_Data_SHINWON(vasIDRow, vasList, Format(dtpExamDate, "yyyymmdd"))
            
            If liRet = 1 Then
                'db_Commit gServer
                
                SetBackColor vasList, vasIDRow, vasIDRow, colBarCode, colState, 255, 255, 180
                SetText vasList, "Trans", vasIDRow, colState
            Else
                SetBackColor vasList, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasList, "Failed", vasIDRow, colState
            End If
            vasList.Col = 1
            vasList.Row = vasIDRow
            vasList.Value = 0
        Else
        
        End If
    Next vasIDRow
    
End Sub

Private Sub cmdLogin_Click()
    Dim strID As String
    
    strID = InputBox("사번입력", "사번입력")
    
    If strID = "사번을 입력해주세요." Or strID = "" Then
        MsgBox "사번을 확인해 주세요."
        Exit Sub
    End If
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT USERID , USERNAME FROM CSUSERMT"
    SQL = SQL & vbCrLf & " WHERE USERID = '" & strID & "' "
    res = db_select_Col(gServer, SQL)
    If res > 0 Then
        sspUsername.Caption = Trim(gReadBuf(1))
        lblSendID.Caption = Trim(gReadBuf(0))
        cmdListTrans.Enabled = True
    Else
        MsgBox "사번을 확인해 주세요."
        sspUsername.Caption = ""
        lblSendID.Caption = ""
        cmdListTrans.Enabled = False
    End If
End Sub

Private Sub cmdResSave_Click()
    '
    Dim vasIDRow As Integer
    Dim VasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer
    Dim iNumber As Integer

    If MsgBox(" " & vbCrLf & "결과를 저장 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:결과저장") = vbCancel Then
        Exit Sub
    End If
    
    For iRow = 1 To vasLISTRES.DataRowCnt
        
        SQL = ""
        SQL = SQL & vbCrLf & "UPDATE PAT_RES"
        SQL = SQL & vbCrLf & "   SET resvalue = '" & Trim(GetText(vasLISTRES, iRow, colResValue)) & "'"
        SQL = SQL & vbCrLf & " WHERE BARCODE  = '" & Trim(Mid(lblInfo1, 1, InStr(1, lblInfo1, "/") - 1)) & "'"
        SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasLISTRES, iRow, colEquipExam)) & "'"
        SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(vasLISTRES, iRow, colExamCode)) & "'"
        If Trim(lblInfo2) <> "" Then
            SQL = SQL & vbCrLf & "   AND TRIM(RECENO) = '20" & Trim(lblInfo2) & "'"
        Else
            SQL = SQL & vbCrLf & "   AND TRIM(RECENO) = '" & Trim(lblInfo2) & "'"
        End If
        res = SendQuery(gLocal, SQL)
        
    Next iRow
    
    
End Sub

Private Sub cmdVasIDWidth_Click()
    Dim i As Integer
    
    
    If cmdVasIDWidth.Caption = ">>" Then
        vasID.Width = Frame3.Width - 180
        cmdVasIDWidth.Caption = "<<"
        
        vasID.Visible = False
        For i = colRStart + 1 To vasID.MaxCols
            vasID.Col = i
            vasID.ColHidden = False
        Next
        vasID.Visible = True
        vasID.ScrollBars = ScrollBarsBoth
    Else
        vasID.Width = 6375
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
        vasList.Width = 6375
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

Public Sub CobasProg(ArgData As String)
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
    
    Dim sInfoValue As String
    Dim sResultValue As String
    Dim sConvertValue As String
    
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
    Dim lsOrderFlag As String
       
    
    Dim sExamCode_All As String
    Dim sPart_All As String
    Dim sBarCode As String
    Dim sBarCode1 As String
    Dim sBarcode2 As String
    Dim sOrDate As String
    
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
    
    'Micro 전송유무 확인 Flag
    Dim strABCDZ_Flag As String
     
    On Error GoTo ErrRes:
    
    Select Case Mid(ArgData, 1, 1)
    Case "H"    'Header
        gCmtFlag = ""
        gPreRow = -1
        strTransYN = "Y"
        
        '오더타입 기억
        strORDERType = ""
        
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
        
        For i = 1 To Len(ArgData)
            If Mid(ArgData, i, 1) = "|" Then
                iCnt = iCnt + 1
                Select Case iCnt
                Case 2  'PID
                    slen = InStr(i + 1, ArgData, "|")
                    sPID = Mid(ArgData, i + 1, slen - i - 1)
                    sSpecID = sPID

                        
                    If sSpecID <> gSpecID Then
                        gSpecID = sSpecID
                    End If
                Case 3 'SAMPLE INFO
                    slen = InStr(i + 1, ArgData, "|")
                    
                    
                    sSampleInfo = Mid(ArgData, i + 1, slen - i - 1)
                    
                    slen = InStr(i + 1, ArgData, "^")
                    
                    gSampleInfo = ""
                    sSampleInfo = Mid(ArgData, i + 1, slen - i + 1)
                    If sSampleInfo <> gSampleInfo Then
                        gSampleInfo = sSampleInfo
                    End If
                    
                    
                Case 11
                    slen = InStr(i + 1, ArgData, "|")
                    If slen > 0 Then
                        If Mid(ArgData, i + 1, slen - i - 1) = "Q" Then
                            sSampleType = "Q"
                        Else
                            sSampleType = "P"
                        End If
                    Else
                        sSampleType = "P"
                    End If
                
                Case 14
                    gExamdate = ""
                    slen = InStr(i + 1, ArgData, "|")
                    If slen > 0 Then
                        gExamdate = Mid(Mid(ArgData, i + 1, slen - i - 1), 1, 8)
                        gExamTime = Mid(Mid(ArgData, i + 1, slen - i - 1), 9)
                    Else
                        gExamdate = Format(Date, "yyyymmdd")
                    End If
                
                End Select
            End If

        Next i
        If sSampleType = "P" Then
        
            glRow = -1
            For lRow = 1 To vasID.DataRowCnt
                If Trim(GetText(vasID, lRow, colBarCode)) = gSpecID Then
                    glRow = lRow
                    
                    If gPatFlag = -1 Then
                        vasID_Click 2, glRow
                        
                        gPatFlag = 1
                        vasActiveCell vasID, glRow, 2
                    End If
    
                    Exit For
                End If
            Next lRow
            
            
            If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
                
                glRow = vasID.DataRowCnt + 1
                
                '새바코드가 나왔을때 위의 줄에서 같은 랙번호의 번호를 지워준다.
                For i = glRow To 1 Step -1
                    If Trim(GetText(vasID, i, colRackPos)) = gSampleInfo And Trim(GetText(vasID, i, colState)) = "오더없음" Then
                        DeleteRow vasID, i, i
                        glRow = glRow - 1
                    End If
                Next i
                
                If glRow > vasID.MaxRows Then
                    vasID.MaxRows = glRow + 1
                End If
                vasActiveCell vasID, glRow, colBarCode
                SetText vasID, sSpecID, glRow, colBarCode
                

                
                
            End If
            
            If Trim(GetText(vasID, glRow, colPID)) = "" Then
                
                Get_Sample_Info glRow
            End If
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
        
        SetText vasID, gSampleInfo, glRow, colRackPos
        
        gPreSpecID = sSpecID
        
        gPreRow = glRow

    Case "R"    'Result
'''        gRecodeType = "R"
        gtestid = ""
        gResultRes = ""
        gCmtFlag = "R"
        
        If Trim(GetText(vasID, glRow, colState)) <> "미접수" And Trim(GetText(vasID, glRow, colState)) <> "오더없음" Then
            SetText vasID, "Result", glRow, colState
        End If
        
        iCnt = 0
        sExamCode = ""
        sResClassCode = ""
        sExamName = ""
        sResult = ""
        
        aCnt = 0
        Dim arrTemp
        Dim arrTemp2
        
        arrTemp = Split(ArgData, "|")
        
        If arrTemp(1) = "1" And arrTemp(13) = "u601" Then
            '오더타입 기억(DISKNO)
            SQL = ""
            SQL = SQL & vbCrLf & "SELECT DISKNO"
            SQL = SQL & vbCrLf & "  FROM PAT_RES "
            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
            SQL = SQL & vbCrLf & "   AND DISKNO <> ''"
            res = db_select_Col(gLocal, SQL)
            
            strORDERType = Trim(gReadBuf(0))
            
            
            '601의 첫번째 결과가 들어올 경우 해당 바코드의 결과를 삭제한다.
            '701결과를 위해 삭제함.
            '자동전송관련
            SQL = ""
            SQL = SQL & vbCrLf & "DELETE FROM PAT_RES"
            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
            'SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & gExamdate & "'"
            'SQL = SQL & vbCrLf & "   AND EXAMTIME = '" & gExamTime & "'"
            SQL = SQL & vbCrLf & "   AND RESVALUE <> ''"
            res = SendQuery(gLocal, SQL)
        End If
        
        '장비코드 확인
        arrTemp2 = Split(arrTemp(2), "^")
        
        
        
        
        sTestID = arrTemp2(1)
        gtestid = sTestID
        
        Dim strExamType As String
        
        SQL = "select examcode, examname, seqno, deltavalue from equipexam " & vbCrLf & _
              "where ExamCode = '" & Trim(gReadBuf(0)) & "' and EquipCode = '" & gtestid & "' "
        SQL = SQL & vbCrLf & "GROUP BY examcode, examname, seqno,  deltavalue"
        SQL = SQL & vbCrLf & " ORDER BY SEQNO DESC"
        res = db_select_Col(gLocal, SQL)
        
        strExamType = Trim(gReadBuf(3))
        
        '결과 확인
        arrTemp2 = Split(arrTemp(3), "^")
        sResValue = arrTemp(3)
        sInfoValue = ""
        sResultValue = ""
        If UBound(arrTemp2) > 0 Then
            If gtestid = "RBC" Or gtestid = "WBC" Then
                sInfoValue = arrTemp2(1)
            ElseIf gtestid = "NEC" Or gtestid = "SEC" Or gtestid = "HYA" Then
                sInfoValue = Replace(arrTemp2(3), "~", "")
            Else
                sInfoValue = arrTemp2(0)
            End If
            
            
            
            sConvertValue = arrTemp2(1)
            sConvertValue = Replace(sConvertValue, " /uL", "")
            sConvertValue = Replace(sConvertValue, " mg/dL", "")
            sConvertValue = Replace(sConvertValue, " mg/mL", "")
            sConvertValue = Replace(sConvertValue, " /HPF", "")
            
            sResultValue = arrTemp2(2)
            sResultValue = Replace(sResultValue, " /uL", "")
            sResultValue = Replace(sResultValue, " mg/dL", "")
            sResultValue = Replace(sResultValue, " mg/mL", "")
            sResultValue = Replace(sResultValue, " /HPF", "")
            sInfoValue = Replace(sInfoValue, " /uL", "")
            sInfoValue = Replace(sInfoValue, " mg/dL", "")
            sInfoValue = Replace(sInfoValue, " mg/mL", "")
            sInfoValue = Replace(sInfoValue, " /HPF", "")
        End If
        
        '결과 확인후 오더 조회
        SQL = "select examcode, equipno from pat_res where barcode = '" & Trim(GetText(vasID, glRow, colBarCode)) & "' and equipcode = '" & gtestid & "'"
        res = db_select_Col(gLocal, SQL)
        
        Dim strExamCodeSRC As String
        Dim intExamCodeSRC As Integer
        strExamCodeSRC = "''"
        
        'PAT_RES에 저장된 오더가 없을경우 서버에서 한번 조회후 오더정보를 확인한다.
        If res < 1 And IsNumeric(gSpecID) = True Then
            SQL = "SELECT EXAMCODE FROM EQUIPEXAM " & vbCrLf & _
                  " WHERE equipno = '" & gEquip & "' " & vbCrLf & _
                  "   and EquipCode = '" & gtestid & "' "
            res = db_select_Row(gLocal, SQL)
            
            For intExamCodeSRC = 0 To res - 1
                If strExamCodeSRC = "''" Then
                    strExamCodeSRC = "'" & gReadBuf(intExamCodeSRC) & "'"
                Else
                    strExamCodeSRC = strExamCodeSRC & ",'" & gReadBuf(intExamCodeSRC) & "'"
                End If
            Next intExamCodeSRC

        End If
        
        '위에서 검사코드가 없을경우 그냥 아무코드나 등록함
        If res = 0 Then
            SQL = "SELECT EXAMCODE FROM EQUIPEXAM " & vbCrLf & _
                  " WHERE equipno = '" & gEquip & "' " & vbCrLf & _
                  "   and EquipCode = '" & gtestid & "' "
            res = db_select_Col(gLocal, SQL)
        End If
        
        '장비코드가 등록되지 않는 검사의 경우 저장하지 않음
        If res = 0 Then Exit Sub
        
        SQL = "select examcode, examname, seqno, deltavalue from equipexam " & vbCrLf & _
              "where ExamCode = '" & Trim(gReadBuf(0)) & "' and EquipCode = '" & gtestid & "' "
        SQL = SQL & vbCrLf & "GROUP BY examcode, examname, seqno,  deltavalue"
        SQL = SQL & vbCrLf & " ORDER BY SEQNO DESC"
        res = db_select_Col(gLocal, SQL)
        
        sExamCode = Trim(gReadBuf(0))
        sExamName = Trim(gReadBuf(1))
        sSeq = Trim(gReadBuf(2))
        
        '결과를 어떤쪽을 쓸껀지 판별
        '현재 코드설정 테이블에서 deltavalue 쪽을 사용함
        '값이 1일 경우 INFO 값을 사용
        '없거나 0일경우 RESULT 값을 사용
        If Trim(GetText(vasID, glRow, colState)) <> "미접수" And Trim(GetText(vasID, glRow, colState)) <> "오더없음" Then
            If Trim(gReadBuf(3)) = "" Or Trim(gReadBuf(3)) = "0" Then
                
                'If gtestid = "NEC" Or gtestid = "SEC" Then
                '    sResValue = sInfoValue
                '    gResultRes = sResValue
                'Else
                    sResValue = sInfoValue
                    gResultRes = sResValue
                'End If
                
                SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 100, 255, 150
                
                '/오더를 체크하여 검사결과가 나오지 않은 항목을 표시해줌
                SQL = ""
                SQL = SQL & vbCrLf & "SELECT URINE, MICRO "
                SQL = SQL & vbCrLf & "  FROM EXAMCHECK"
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                res = db_select_Col(gLocal, SQL)
                
                If Trim(gReadBuf(0)) = "Y" And Trim(gReadBuf(1)) = "Y" Then
                    SQL = ""
                    SQL = SQL & vbCrLf & "UPDATE EXAMCHECK"
                    SQL = SQL & vbCrLf & "   SET MICRO = ''"
                    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                    res = SendQuery(gLocal, SQL)
                    
                ElseIf Trim(gReadBuf(0)) = "" And Trim(gReadBuf(1)) = "Y" Then
                    SQL = ""
                    SQL = SQL & vbCrLf & "DELETE FROM EXAMCHECK"
                    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                    res = SendQuery(gLocal, SQL)
                End If
                
            ElseIf Trim(gReadBuf(3)) = "1" Then
                If gtestid = "UBG" Then
                    sResValue = sInfoValue
                    gResultRes = sResValue
                Else
                    sResValue = sInfoValue
                    gResultRes = sResValue
                End If
                SetBackColor vasID, glRow, glRow, colOrdStick, colOrdStick, 100, 255, 150
                
                '/오더를 체크하여 검사결과가 나오지 않은 항목을 표시해줌
                SQL = ""
                SQL = SQL & vbCrLf & "SELECT URINE, MICRO "
                SQL = SQL & vbCrLf & "  FROM EXAMCHECK"
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                res = db_select_Col(gLocal, SQL)
                
                If Trim(gReadBuf(0)) = "Y" And Trim(gReadBuf(1)) = "Y" Then
                    SQL = ""
                    SQL = SQL & vbCrLf & "UPDATE EXAMCHECK"
                    SQL = SQL & vbCrLf & "   SET URINE = ''"
                    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                    res = SendQuery(gLocal, SQL)
                    
                ElseIf Trim(gReadBuf(0)) = "Y" And Trim(gReadBuf(1)) = "" Then
                    SQL = ""
                    SQL = SQL & vbCrLf & "DELETE FROM EXAMCHECK"
                    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                    res = SendQuery(gLocal, SQL)
                End If
                
            End If
        Else
            If Trim(gReadBuf(3)) = "" Or Trim(gReadBuf(3)) = "0" Then
                sResValue = sResultValue
                gResultRes = sResValue
                SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 50, 50, 255
                
                '/오더를 체크하여 검사결과가 나오지 않은 항목을 표시해줌
                SQL = ""
                SQL = SQL & vbCrLf & "SELECT URINE, MICRO "
                SQL = SQL & vbCrLf & "  FROM EXAMCHECK"
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                res = db_select_Col(gLocal, SQL)
                
                If Trim(gReadBuf(0)) = "Y" And Trim(gReadBuf(1)) = "Y" Then
                    SQL = ""
                    SQL = SQL & vbCrLf & "UPDATE EXAMCHECK"
                    SQL = SQL & vbCrLf & "   SET MICRO = ''"
                    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                    res = SendQuery(gLocal, SQL)
                    
                ElseIf Trim(gReadBuf(0)) = "" And Trim(gReadBuf(1)) = "Y" Then
                    SQL = ""
                    SQL = SQL & vbCrLf & "DELETE FROM EXAMCHECK"
                    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                    res = SendQuery(gLocal, SQL)
                End If
                
                
            ElseIf Trim(gReadBuf(3)) = "1" Then
                sResValue = sInfoValue
                gResultRes = sResValue
                SetBackColor vasID, glRow, glRow, colOrdStick, colOrdStick, 50, 50, 255
                
                '/오더를 체크하여 검사결과가 나오지 않은 항목을 표시해줌
                SQL = ""
                SQL = SQL & vbCrLf & "SELECT URINE, MICRO "
                SQL = SQL & vbCrLf & "  FROM EXAMCHECK"
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                res = db_select_Col(gLocal, SQL)
                
                If Trim(gReadBuf(0)) = "Y" And Trim(gReadBuf(1)) = "Y" Then
                    SQL = ""
                    SQL = SQL & vbCrLf & "UPDATE EXAMCHECK"
                    SQL = SQL & vbCrLf & "   SET URINE = ''"
                    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                    res = SendQuery(gLocal, SQL)
                    
                ElseIf Trim(gReadBuf(0)) = "Y" And Trim(gReadBuf(1)) = "" Then
                    SQL = ""
                    SQL = SQL & vbCrLf & "DELETE FROM EXAMCHECK"
                    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                    res = SendQuery(gLocal, SQL)
                End If
                
                
            End If
        End If
        
        If sResValue = "neg" Then
            sResValue = "Neg"
            gResultRes = "Neg"
        ElseIf sResValue = "pos" Then
            sResValue = "Pos"
            gResultRes = "Pos"
        ElseIf sResValue = "found" Then
            sResValue = "Found"
            gResultRes = "Found"
        ElseIf sResValue = "not found" Then
            sResValue = "Not found"
            gResultRes = "Not found"
        End If
        
        sGiho = ""
        If Left(sResValue, 1) = "<" Or Left(sResValue, 1) = ">" Then
            '기호를 없애버림
            'sGiho = Left(sResValue, 1)
            sGiho = ""
            sResValue = Trim(Mid(sResValue, 2))
        End If
        
        '결과 변환
        sResValue = CONVERT_RESULT(gtestid, sResValue)
        '결과 확인
        sResult1 = Result_Set(sExamCode, sResValue, gtestid)
        
        j = InStr(1, sResult1, "|")
        
        sResValue = sGiho & Mid(sResult1, 1, j - 1)
        sResult = Mid(sResult1, j + 1)
        
        sFlag = ""
        '플래그 처리
        sFlag = Result_FLAG(sResValue, gtestid)
                
        
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
                
                lsOrderFlag = ""
                If sExamCode <> "" And IsNumeric(gSpecID) = True And Len(gSpecID) = 12 Then
                    SQL = "SELECT R.ORDCODE "
                    SQL = SQL & vbCrLf & "   FROM SLRSLTMT R"
                    SQL = SQL & vbCrLf & " WHERE 1 = 1"
                    SQL = SQL & vbCrLf & "   AND R.SPCDATE = TO_DATE(SUBSTR('" & gSpecID & "', 1, 6), 'YYMMDD')"
                    SQL = SQL & vbCrLf & "   AND R.SPCNO = SUBSTR('" & gSpecID & "', 7, 5)"
                    SQL = SQL & vbCrLf & "   AND R.SPCSEQ = SUBSTR('" & gSpecID & "', 12, 1)"
                    'SQL = SQL & vbCrLf & "   AND R.PROCSTAT IN ('D', 'E')"
                    SQL = SQL & vbCrLf & "   AND R.ORDCODE  = '" & sExamCode & "'"
                    SQL = SQL & vbCrLf & " GROUP BY R.ORDCODE "
                    
'''
'''                    SQL = ""
'''                    SQL = SQL & vbCrLf & "SELECT CASE WHEN TESTSUBCODE <> '' THEN '*'"
'''                    SQL = SQL & vbCrLf & "       END"
'''                    SQL = SQL & vbCrLf & "  FROM LabRegResult "
'''                    SQL = SQL & vbCrLf & " WHERE LabRegDate = '" & "20" & Format(Mid(gSpecID, 1, 6), "@@-@@-@@") & "'"
'''                    SQL = SQL & vbCrLf & "   AND LabRegNo = '" & Trim(Mid(gSpecID, 7)) & "'"
'''                    SQL = SQL & vbCrLf & "   AND TESTSUBCODE = '" & sExamCode & "'"
                    res = db_select_Col(gServer, SQL)
                    
                    lsOrderFlag = Trim(gReadBuf(0))
                    If lsOrderFlag = "" Then
                        lsOrderFlag = "*"
                    Else
                        lsOrderFlag = ""
                    End If
                Else
                    lsOrderFlag = "*"
                End If
                
                If sExamDate = "" Then sExamDate = gExamdate
                
                
                SetText vasRes, gtestid, lResRow, colEquipExam '장비코드
                SetText vasRes, sExamCode, lResRow, colExamCode '검사코드
                SetText vasRes, sExamName, lResRow, colExamName '검사명
                SetText vasRes, sSeq, lResRow, colSeq '순서
                SetText vasRes, sResValue, lResRow, colResValue '결과수치
                SetText vasRes, sResult, lResRow, colResult '문자결과
                SetText vasRes, gExamdate, lResRow, colResDate '검사일자
                SetText vasRes, gExamTime, lResRow, colResTime '검사시간
                SetText vasRes, sFlag, lResRow, colAFLAG 'HL
                SetText vasRes, lsOrderFlag, lResRow, colORDERFLAG '검사시간
                
                Save_Local_One glRow, lResRow, "1"
                    
                'vasID_Click colBarCode, glRow
                
                SQL = "select resgubun from equipexam " & vbCrLf & _
                      "where equipno = '" & gEquip & "' and examcode = '" & sExamCode & "'"
                res = db_select_Col(gLocal, SQL)
                
                If Trim(gReadBuf(0)) = "1" Then
                    SetPositionResult glRow, gtestid, sResult & "(" & sResValue & ")"
                Else
                
                    SetPositionResult glRow, gtestid, sResValue
                End If

            ElseIf sSampleType = "Q" Then   'QC결과
                
            End If
        End If
        
        gMsgFlag = ""
        gHeadRecode = ""
        txtData.Text = ""
        
    Case "Q"    'Request
        gRecodeType = "Q"
        
        ClearSpread vasTemp
        ClearSpread vasOrder
        ClearSpread vasOrderBuf
        
        slen = InStr(1, ArgData, "|")
        ArgData = Mid(ArgData, slen + 1)
        
        slen = InStr(1, ArgData, "|")
        ArgData = Mid(ArgData, slen + 1)
        
        slen = InStr(1, ArgData, "^")
        ArgData = Mid(ArgData, slen + 1)     '검체번호
        
        slen = InStr(1, ArgData, "^")
        gSpecID = Mid(ArgData, 1, slen - 1)    '검체번호
        'gSpecID = Mid(gSpecID, 3)
        
        slen = InStr(1, ArgData, "^")
        
        gSampleInfo = ""
        sSampleInfo = Trim(Mid(ArgData, slen + 1))         'sampleinfo

        gSampleInfo = sSampleInfo
        
'''        glRow = vasID.DataRowCnt + 1
'''        If vasID.MaxRows <= glRow Then
'''            vasID.MaxRows = glRow + 1
'''        End If
        
        glRow = -1
        For lRow = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, lRow, colBarCode)) = gSpecID Then
                glRow = lRow

                Exit For
            End If
        Next lRow
        
        '==========================================================================
        'Order 전송뒤 Clear시 다시 바코드 스캔 안하고 결과 넘어오도록 수정
        If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
            glRow = vasID.DataRowCnt + 1
            If glRow > vasID.MaxRows Then
                vasID.MaxRows = glRow + 1
            End If
            
            vasActiveCell vasID, glRow, colBarCode
            SetText vasID, gSpecID, glRow, colBarCode
        End If
        '==========================================================================
        If Trim(GetText(vasID, glRow, colPID)) = "" Then
            Get_Sample_Info glRow
        End If
        
        SetText vasID, gSampleInfo, glRow, colRackPos
        
        'Order 만들기
        Make_Order gSpecID, glRow, gSampleInfo
        Save_Raw_Data "[TX " & Format(Time, "hh:mm:ss") & "] " & "오더처리완료"
    Case "L"    '자료수신 완료
        If gRecodeType = "Q" Then
        Else
            
            If gCmtFlag <> "R" Then
                '검사에러일때 결과가 없는 경우가 발생.
                '더이상 진행하지 않고 처리 한다.
                Exit Sub
            End If
            'vasID_Click colBarCode, glRow
            
            
            '기본값 전송이 없어짐.(나중에 생길수도 있음)
'''            Dim intMicro As Integer
'''
'''            'STICK 검사결과에 이상이 없으면 Micro 기본값 넣어줌
'''            SQL = ""
'''            SQL = SQL & vbCrLf & "SELECT RESVALUE"
'''            SQL = SQL & vbCrLf & "  FROM PAT_RES"
'''            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, glRow, colBarCode)) & "'"
'''            SQL = SQL & vbCrLf & "   AND RESVALUE <> ''"
'''            SQL = SQL & vbCrLf & "   AND EQUIPCODE = 'ERY' "
'''            SQL = SQL & vbCrLf & "UNION ALL"
'''            SQL = SQL & vbCrLf & "SELECT RESVALUE"
'''            SQL = SQL & vbCrLf & "  FROM PAT_RES"
'''            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, glRow, colBarCode)) & "'"
'''            SQL = SQL & vbCrLf & "   AND RESVALUE <> ''"
'''            SQL = SQL & vbCrLf & "   AND EQUIPCODE = 'LEU' "
'''            res = db_select_Row(gLocal, SQL)
'''
'''            If res = 0 Then
'''                strTransYN = "N"
'''            ElseIf res = 2 Then  'ERY와 LEU 의 결과가 있을때에만 결과를 넣어줌
'''                'STICK 검사결과에 이상이 없으면 Micro 기본값 넣어줌
'''                SQL = ""
'''                SQL = SQL & vbCrLf & "SELECT RESVALUE"
'''                SQL = SQL & vbCrLf & "  FROM PAT_RES"
'''                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, glRow, colBarCode)) & "'"
'''                SQL = SQL & vbCrLf & "   AND RESVALUE IN ('1+','2+','3+','4+','5+','Trace') "
'''                SQL = SQL & vbCrLf & "   AND EQUIPCODE = 'ERY' "
'''                SQL = SQL & vbCrLf & "UNION ALL"
'''                SQL = SQL & vbCrLf & "SELECT RESVALUE"
'''                SQL = SQL & vbCrLf & "  FROM PAT_RES"
'''                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, glRow, colBarCode)) & "'"
'''                SQL = SQL & vbCrLf & "   AND RESVALUE IN ('1+','2+','3+','4+','5+','Trace') "
'''                SQL = SQL & vbCrLf & "   AND EQUIPCODE = 'LEU' "
'''                res = db_select_Row(gLocal, SQL)
'''
'''                If res = 0 Then
'''                    'STICK 오더만 있을경우 MICRO 검사결과도 만들어 준다.
'''                    For intMicro = 1 To 7
'''                        SQL = ""
'''                        SQL = SQL & vbCrLf & "select EQUIPCODE, examcode, examname, seqno "
'''                        SQL = SQL & vbCrLf & "  FROM EQUIPEXAM "
'''                        Select Case intMicro
'''                            Case 1
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'WBC'"
'''                            Case 2
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'RBC'"
'''                            Case 3
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'EPCEL'"
'''                            Case 4
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'BAC'"
'''                            Case 5
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'CRY'"
'''                            Case 6
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'PAT'"
'''                            Case 7
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'OTR'"
'''                        End Select
'''
'''                        SQL = SQL & vbCrLf & " GROUP BY EQUIPCODE, examcode, examname, seqno"
'''                        SQL = SQL & vbCrLf & " ORDER BY SEQNO DESC"
'''                        res = db_select_Col(gLocal, SQL)
'''                        gtestid = Trim(gReadBuf(0))
'''                        sExamCode = Trim(gReadBuf(1))
'''                        sExamName = Trim(gReadBuf(2))
'''                        sSeq = Trim(gReadBuf(3))
'''
'''                        'PAT_RES에 해당 장비코드로 등록된 레코드가 없을경우 하나 만들어줌
'''                        SQL = "select examcode, equipno from pat_res "
'''                        SQL = SQL & vbCrLf & "where barcode = '" & Trim(GetText(vasID, glRow, colBarCode)) & "' and equipcode = '" & gtestid & "'"
'''                        SQL = SQL & vbCrLf & "   AND RESVALUE <> ''"
'''                        'SQL = SQL & vbCrLf & "  AND POSNO = '" & Trim(GetText(vasID, glRow, colRackPos)) & "'"
'''                        res = db_select_Col(gLocal, SQL)
'''
'''                        If res = 0 Then
'''                            lResRow = -1
'''                            For j = 1 To vasRes.DataRowCnt
'''                                If Trim(sExamCode) = Trim(GetText(vasRes, j, colExamCode)) Then
'''                                    lResRow = j
'''                                    Exit For
'''                                End If
'''                            Next j
'''
'''                            If lResRow = -1 Then
'''                                lResRow = vasRes.DataRowCnt + 1
'''                                If lResRow > vasRes.MaxRows Then
'''                                    vasRes.MaxRows = lResRow
'''                                End If
'''                            End If
'''
'''                            SetText vasRes, gtestid, lResRow, colEquipExam '장비코드
'''                            SetText vasRes, sExamCode, lResRow, colExamCode '검사코드
'''                            SetText vasRes, sExamName, lResRow, colExamName '검사명
'''                            SetText vasRes, sSeq, lResRow, colSeq '순서
'''                            SetText vasRes, "", lResRow, colResValue '결과수치
'''                            SetText vasRes, "", lResRow, colResult '문자결과
'''                            SetText vasRes, Trim(GetText(vasRes, 1, colResDate)), lResRow, colResDate '검사일자
'''                            SetText vasRes, Trim(GetText(vasRes, 1, colResTime)), lResRow, colResTime '검사시간
'''
'''                            SetText vasRes, "*", lResRow, colORDERFLAG '검사시간
'''
'''                            Save_Local_One glRow, lResRow, "1"
'''                        End If
'''
'''                    Next intMicro
'''
'''                    '뇨침사 결과 기본값 저장
'''                    '스틱검사결과가 들어온 후 뇨침사 결과가 들어오는데 스틱 결과가 들어온후에 기본결과를 자동으로 저장한다.
'''                    '뇨침사 결과가 들어올 경우 바로 결과를 업데이트 한다.
'''                    '7가지 항목이므로 7번 업데이트 작업함.
'''                    For intMicro = 1 To 7
'''                        SQL = ""
'''                        SQL = SQL & vbCrLf & "UPDATE PAT_RES"
'''                        SQL = SQL & vbCrLf & "   SET SENDFLAG = '1'"
'''                        SQL = SQL & vbCrLf & "      ,RECENO = '" & Mid(Trim(GetText(vasID, glRow, colBarCode)), 7, 5) & "'"
'''                        Select Case intMicro
'''                            Case 1
'''                                SQL = SQL & vbCrLf & "      ,RESVALUE = '0 - 1'"
'''                                SQL = SQL & vbCrLf & "      ,POSNO = '" & Trim(GetText(vasID, glRow, colRackPos)) & "'"
'''                                SQL = SQL & vbCrLf & "      ,reFflag = ''"
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'WBC'"
'''                            Case 2
'''                                SQL = SQL & vbCrLf & "      ,RESVALUE = '0 - 1'"
'''                                SQL = SQL & vbCrLf & "      ,POSNO = '" & Trim(GetText(vasID, glRow, colRackPos)) & "'"
'''                                SQL = SQL & vbCrLf & "      ,reFflag = ''"
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'RBC'"
'''                            Case 3
'''                                SQL = SQL & vbCrLf & "      ,RESVALUE = '0 - 1'"
'''                                SQL = SQL & vbCrLf & "      ,POSNO = '" & Trim(GetText(vasID, glRow, colRackPos)) & "'"
'''                                SQL = SQL & vbCrLf & "      ,reFflag = ''"
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'EPCEL'"
'''                            Case 4
'''                                SQL = SQL & vbCrLf & "      ,RESVALUE = 'Not found'"
'''                                SQL = SQL & vbCrLf & "      ,POSNO = '" & Trim(GetText(vasID, glRow, colRackPos)) & "'"
'''                                SQL = SQL & vbCrLf & "      ,reFflag = ''"
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'BAC'"
'''                            Case 5
'''                                SQL = SQL & vbCrLf & "      ,RESVALUE = 'Not found'"
'''                                SQL = SQL & vbCrLf & "      ,POSNO = '" & Trim(GetText(vasID, glRow, colRackPos)) & "'"
'''                                SQL = SQL & vbCrLf & "      ,reFflag = ''"
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'CRY'"
'''                            Case 6
'''                                SQL = SQL & vbCrLf & "      ,RESVALUE = 'Not found'"
'''                                SQL = SQL & vbCrLf & "      ,POSNO = '" & Trim(GetText(vasID, glRow, colRackPos)) & "'"
'''                                SQL = SQL & vbCrLf & "      ,reFflag = ''"
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'PAT'"
'''                            Case 7
'''                                SQL = SQL & vbCrLf & "      ,RESVALUE = 'Not found'"
'''                                SQL = SQL & vbCrLf & "      ,POSNO = '" & Trim(GetText(vasID, glRow, colRackPos)) & "'"
'''                                SQL = SQL & vbCrLf & "      ,reFflag = ''"
'''                                SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'OTR'"
'''                        End Select
'''                        SQL = SQL & vbCrLf & "  AND RESVALUE = ''"
'''                        SQL = SQL & vbCrLf & "  AND BARCODE = '" & Trim(GetText(vasID, glRow, colBarCode)) & "'"
'''                        'SQL = SQL & vbCrLf & "  AND POSNO = '" & Trim(GetText(vasID, glRow, colRackPos)) & "'"
'''                        res = SendQuery(gLocal, SQL)
'''
'''                    Next intMicro
                    
                    If GetText(vasID, glRow, colMicroCheck) <> "" Then
                        SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 100, 255, 150
'
'                        '/오더를 체크하여 검사결과가 나오지 않은 항목을 표시해줌
'                        SQL = ""
'                        SQL = SQL & vbCrLf & "SELECT URINE, MICRO "
'                        SQL = SQL & vbCrLf & "  FROM EXAMCHECK"
'                        SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, glRow, colBarCode)) & "'"
'                        res = db_select_Col(gLocal, SQL)
'
'                        If Trim(gReadBuf(0)) = "Y" And Trim(gReadBuf(1)) = "Y" Then
'                            SQL = ""
'                            SQL = SQL & vbCrLf & "UPDATE EXAMCHECK"
'                            SQL = SQL & vbCrLf & "   SET MICRO = ''"
'                            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, glRow, colBarCode)) & "'"
'                            res = SendQuery(gLocal, SQL)
'
'                        ElseIf Trim(gReadBuf(0)) = "" And Trim(gReadBuf(1)) = "Y" Then
'                            SQL = ""
'                            SQL = SQL & vbCrLf & "DELETE FROM EXAMCHECK"
'                            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, glRow, colBarCode)) & "'"
'                            res = SendQuery(gLocal, SQL)
'                        End If
'
                    End If
            
            '결과 확정 플레그 처리.
            strFSend = ""
            strFSend = RES_MICROCHECK(Trim(GetText(vasID, glRow, colBarCode)))
            'strFSend = RESULT_VALIDATE(Trim(GetText(vasID, glRow, colBarCode)))
            
            
            strABCDZ_Flag = RES_FLAGCHECK(Trim(GetText(vasID, glRow, colBarCode)))
            If strABCDZ_Flag <> "" Then
                SQL = ""
                SQL = SQL & vbCrLf & "UPDATE PAT_RES"
                SQL = SQL & vbCrLf & "   SET RESFLAG = '" & strABCDZ_Flag & "'"
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                res = SendQuery(gLocal, SQL)
                If strABCDZ_Flag <> "" Then
                
                    If strFSend = "YY" And InStr(strABCDZ_Flag, "A") > 0 Then
                        'Micro는 전송하지 않게 변경함.
                        '신장내과가 있는경우.
                        strFSend = "Y"
                    ElseIf InStr(strABCDZ_Flag, "A") > 0 Then
                        'Micro는 전송하지 않게 변경함.
                        '신장내과가 있는경우.
                        strFSend = "Y"
                    End If
                    
                    
                    
                    SetVasColor vasID, CInt(glRow), Trim(GetText(vasID, glRow, colBarCode))
                    
                End If
            End If
            SetForeColor vasID, glRow, glRow, colOrdMicroYN, colOrdMicroYN, 255, 255, 255
            vasID.Row = glRow
            vasID.Row = colOrdMicroYN
            vasID.RowHeight(glRow) = 21.75
            'vasID.FontSize = 13
            vasID.FontBold = True
            If strFSend = "NN" Then
            'Micro 결과가 비정상일경우에 MicroYN 에 Y를 표시해준다.
                SQL = ""
                SQL = SQL & vbCrLf & "UPDATE PAT_RES"
                SQL = SQL & vbCrLf & "   SET EXAMTYPE = 'Y'"
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                res = SendQuery(gLocal, SQL)
                SetText vasID, "Y", glRow, colOrdMicroYN
                SetForeColor vasID, glRow, glRow, colOrdMicroYN, colOrdMicroYN, 255, 50, 50
                strFSend = "Y"
            ElseIf strFSend = "YY" Then
                SQL = ""
                SQL = SQL & vbCrLf & "UPDATE PAT_RES"
                SQL = SQL & vbCrLf & "   SET EXAMTYPE = 'N'"
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                res = SendQuery(gLocal, SQL)
                'SetText vasID, "N", glRow, colOrdMicroYN
            ElseIf strFSend = "U" Then
                SQL = ""
                SQL = SQL & vbCrLf & "UPDATE PAT_RES"
                SQL = SQL & vbCrLf & "   SET EXAMTYPE = 'U'"
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
                res = SendQuery(gLocal, SQL)
                SetText vasID, "U", glRow, colOrdMicroYN
                SetForeColor vasID, glRow, glRow, colOrdMicroYN, colOrdMicroYN, 255, 50, 50
                strFSend = "Y"
            End If
            
            'vasID_Click colBarCode, glRow
            If mnuAuto.Checked = True Then
                If glRow > 0 And glRow <= vasID.DataRowCnt And IsNumeric(gSpecID) = True Then
                    res = -1
                    res = Insert_Data(CInt(glRow), vasID, Trim(GetText(vasRes, 1, colResDate)), strFSend, "Y")
                    If res = 1 Then
                        SQL = ""
                        SQL = SQL & vbCrLf & "SELECT SENDFLAG "
                        SQL = SQL & vbCrLf & "  FROM PAT_RES"
                        SQL = SQL & vbCrLf & " WHERE BARCODE = '"
'                            SetBackColor vasID, gPreRow, gPreRow, colCheckBox, colState, 202, 255, 112
                        'SetBackColor vasID, glRow, glRow, 2, colPSex, 220, 250, 220
                        SetBackColor vasID, glRow, glRow, colState, colState, 220, 250, 220
                        If Trim(GetText(vasID, glRow, colState)) <> "미접수" And Trim(GetText(vasID, glRow, colState)) <> "오더없음" Then
                            SetText vasID, "Trans", glRow, colState
                        End If
                    ElseIf res = -1 Then
                        SetBackColor vasID, glRow, glRow, 2, colPSex, 255, 0, 0
                        SetBackColor vasID, glRow, glRow, colState, colState, 255, 0, 0
                        If Trim(GetText(vasID, glRow, colState)) <> "미접수" And Trim(GetText(vasID, glRow, colState)) <> "오더없음" Then
                            SetText vasID, "Failed", glRow, colState
                        End If
                    End If
                End If
            End If
        End If
        
        If strQMode <> "Q" Then
            Save_Raw_Data "[TX " & Format(Time, "hh:mm:ss") & "] " & "결과 처리완료"
        End If
        
        
    End Select
    
    
    Exit Sub
    
ErrRes:
    Save_Raw_Data "결과처리 에러-" & Format(Now, "hhmmss") & "-" & gSpecID
    Exit Sub
    
End Sub

Function RES_FLAGCHECK(argBarcode As String) As String
    RES_FLAGCHECK = ""
    Dim i As Integer
    
    Dim strAGE As String
    Dim strDept As String
    
    'NEMD 환자
    Dim strFlag_A As String
    'Critical V
    Dim strFlag_B As String
    '비중측정오류
    Dim strFlag_C As String
    'Glucose정량
    Dim strFlag_D As String
    '12세미만
    Dim strFlag_Z As String
    
    'B의 경우 3가지 조건 만족할때 플래그처리한다.
    Dim intFlag_B As Integer
    'D의 경우 5가지 조건을 만족할때 프래그처리한다.
    Dim intFlag_D As Integer
    
    
    
    SQL = "    SELECT DISTINCT FN_SD_GET_AGE(TO_CHAR(SYSDATE,'YYYY-MM-DD'), P.BIRTHDAY )"
    SQL = SQL & vbCrLf & "    ,(SELECT DISTINCT MEDDEPT FROM SLACPTMT WHERE ROWNUM = 1 AND SPCDATE = R.SPCDATE AND SPCNO = R.SPCNO AND SPCSEQ = R.SPCSEQ )"
    SQL = SQL & vbCrLf & "  FROM SLRSLTMT R, ACPATBAT P"
    SQL = SQL & vbCrLf & " WHERE R.SPCDATE = TO_DATE(SUBSTR('" & argBarcode & "', 1, 6), 'YYMMDD')"
    SQL = SQL & vbCrLf & "   AND R.SPCNO = SUBSTR('" & argBarcode & "', 7, 5)"
    SQL = SQL & vbCrLf & "   AND R.SPCSEQ = SUBSTR('" & argBarcode & "', 12, 1)"
    SQL = SQL & vbCrLf & "   AND R.PATNO = P.PATNO"
    res = db_select_Col(gServer, SQL)
    
    strAGE = Replace(Trim(gReadBuf(0)), "세", "")
    strDept = Trim(gReadBuf(1))
    
    If InStr(1, strDept, "NEMD") > 0 Then
        strFlag_A = "A"
    End If
    
    If IsNumeric(strAGE) = True Then
        If CCur(strAGE) < 12 Then
            strFlag_Z = "Z"
        End If
    Else
        '나이가 문자일경우 아기로 판단한다.
        strFlag_Z = "Z"
    End If
    
    intFlag_B = 0
    intFlag_D = 0
    
    ClearSpread vasListTemp
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT EQUIPCODE, RESVALUE"
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & argBarcode & "'"
    res = db_select_Vas(gLocal, SQL, vasListTemp)
    
    For i = 1 To vasListTemp.DataRowCnt
        If Trim(GetText(vasListTemp, i, 1)) = "KET" Then
            If Trim(GetText(vasListTemp, i, 2)) = "1+" Then
                intFlag_D = intFlag_D + 1
                
            ElseIf Trim(GetText(vasListTemp, i, 2)) = "2+" Then
                intFlag_D = intFlag_D + 1
                
            ElseIf Trim(GetText(vasListTemp, i, 2)) = "3+" Then
                intFlag_D = intFlag_D + 1
                intFlag_B = intFlag_B + 1
                
            ElseIf Trim(GetText(vasListTemp, i, 2)) = "4+" Then
                intFlag_D = intFlag_D + 1
                intFlag_B = intFlag_B + 1
                
            End If
            
        End If
        
        If Trim(GetText(vasListTemp, i, 1)) = "SG" Then
            If IsNumeric(Trim(GetText(vasListTemp, i, 2))) = True Then
                If CCur(Trim(GetText(vasListTemp, i, 2))) >= 1.03 Then
                    intFlag_D = intFlag_D + 1
                End If
            Else
                If Trim(GetText(vasListTemp, i, 2)) = "-" Then
                    strFlag_C = "C"
                End If
            End If
        End If
        
        If Trim(GetText(vasListTemp, i, 1)) = "pH" Then
            If IsNumeric(Trim(GetText(vasListTemp, i, 2))) = True Then
                If CCur(Trim(GetText(vasListTemp, i, 2))) >= 6 Then
                    intFlag_D = intFlag_D + 1
                End If
            
            End If
        End If
        
        If Trim(GetText(vasListTemp, i, 1)) = "GLU" Then
            If Trim(GetText(vasListTemp, i, 2)) = "Neg" Then
                intFlag_D = intFlag_D + 1
            ElseIf Trim(GetText(vasListTemp, i, 2)) = "3+" Then
                intFlag_B = intFlag_B + 1
            End If
        End If
        
    Next i
    
    'FLAG 만족일때
    If intFlag_B = 2 And strFlag_Z = "Z" Then
        strFlag_B = "B"
    End If
    
    If intFlag_D = 4 And strFlag_Z = "Z" Then
        strFlag_D = "D"
    End If
    
    RES_FLAGCHECK = strFlag_A & strFlag_B & strFlag_C & strFlag_D & strFlag_Z
    
End Function

Function RES_MICROCHECK(argBarcode As String) As String
    '이상결과가 있는지 없는지 확인한다.
    RES_MICROCHECK = "N"
    Dim i As Integer
    Dim strRBC As String
    Dim strWBC As String
    Dim strBAC As String
    Dim strYEA As String
    Dim strCRY As String
    Dim strPAT As String
    Dim strNEC As String
    Dim strSEC As String
    Dim strHYA As String
    Dim strMUC As String
    
    Dim strCheck As String
    Dim strMCheck As String
    strCheck = "Y"
    
    ClearSpread vasListTemp
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT EQUIPCODE, RESVALUE"
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & argBarcode & "'"
    res = db_select_Vas(gLocal, SQL, vasListTemp)
    
    
    For i = 1 To vasListTemp.DataRowCnt
        strCheck = RES_CHECK(Trim(GetText(vasListTemp, i, 1)), UCase(Trim(GetText(vasListTemp, i, 2))))
        '이상결과가 있을때 빠져나간다.
        If strCheck = "N" Then Exit For
    Next i
    
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT COUNT(A.EQUIPCODE)"
    SQL = SQL & vbCrLf & "  FROM PAT_RES A, EQUIPEXAM B"
    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & argBarcode & "'"
    SQL = SQL & vbCrLf & "   AND A.EQUIPCODE = B.EQUIPCODE"
    SQL = SQL & vbCrLf & "   AND B.DELTAVALUE = '0'"
    SQL = SQL & vbCrLf & "   AND A.RESVALUE <> ''"
    res = db_select_Col(gLocal, SQL)
    
    If gReadBuf(0) = "0" Then
        strMCheck = ""
    Else
        strMCheck = "Y"
        If strCheck = "Y" Then
            For i = 1 To vasListTemp.DataRowCnt
                If Trim(GetText(vasListTemp, i, 1)) = "RBC" Then
                    If IsNumeric(Trim(GetText(vasListTemp, i, 2))) = True Then
                        If CCur(Trim(GetText(vasListTemp, i, 2))) >= 5 Then
                            strMCheck = "N"
                            Exit For
                        End If
                    Else
                        strMCheck = "N"
                        Exit For
                    End If
                ElseIf Trim(GetText(vasListTemp, i, 1)) = "WBC" Then
                    If IsNumeric(Trim(GetText(vasListTemp, i, 2))) = True Then
                        If CCur(Trim(GetText(vasListTemp, i, 2))) >= 5 Then
                            strMCheck = "N"
                            Exit For
                        End If
                    Else
                        strMCheck = "N"
                        Exit For
                    End If
                ElseIf Trim(GetText(vasListTemp, i, 1)) = "BAC" Then
                    If Trim(GetText(vasListTemp, i, 2)) <> "Neg" Then
                        strMCheck = "N"
                        Exit For
                    End If
                ElseIf Trim(GetText(vasListTemp, i, 1)) = "YEA" Then
                    If Trim(GetText(vasListTemp, i, 2)) <> "Neg" Then
                        strMCheck = "N"
                        Exit For
                    End If
                ElseIf Trim(GetText(vasListTemp, i, 1)) = "CRY" Then
                    If Trim(GetText(vasListTemp, i, 2)) <> "Neg" Then
                        strMCheck = "N"
                        Exit For
                    End If
                ElseIf Trim(GetText(vasListTemp, i, 1)) = "PAT" Then
                    If Trim(GetText(vasListTemp, i, 2)) <> "Neg" Then
                        strMCheck = "N"
                        Exit For
                    End If
                ElseIf Trim(GetText(vasListTemp, i, 1)) = "NEC" Then
                    If IsNumeric(Trim(GetText(vasListTemp, i, 2))) = True Then
                        If CCur(Trim(GetText(vasListTemp, i, 2))) >= 0.24 Then
                            strMCheck = "N"
                            Exit For
                        End If
                    Else
                        strMCheck = "N"
                        Exit For
                    End If
                    
                ElseIf Trim(GetText(vasListTemp, i, 1)) = "SEC" Then
                    If IsNumeric(Trim(GetText(vasListTemp, i, 2))) = True Then
                        If CCur(Trim(GetText(vasListTemp, i, 2))) >= 1.15 Then
                            strMCheck = "N"
                            Exit For
                        End If
                    Else
                        strMCheck = "N"
                        Exit For
                    End If
                ElseIf Trim(GetText(vasListTemp, i, 1)) = "HYA" Then
                    If IsNumeric(Trim(GetText(vasListTemp, i, 2))) = True Then
                        If CCur(Trim(GetText(vasListTemp, i, 2))) >= 0.5 Then
                            strMCheck = "N"
                            Exit For
                        End If
                    Else
                        strMCheck = "N"
                        Exit For
                    End If
                ElseIf Trim(GetText(vasListTemp, i, 1)) = "MUC" Then
                    If Trim(GetText(vasListTemp, i, 2)) <> "Neg" Then
                        strMCheck = "N"
                        Exit For
                    End If
                
                End If
            Next i
        End If
    End If
    
    
    If strMCheck = "Y" And strCheck = "Y" Then
        'MICRO결과가 정상적일 경우 YY를 만들어줌
        strCheck = "YY"
    ElseIf strMCheck = "" And strCheck = "Y" Then
        'MICRO결과가 없을경우
        strCheck = "YN"
    ElseIf strMCheck <> "" And strCheck = "N" Then
        'MICRO결과가 없을경우
        strCheck = "N"
    ElseIf strMCheck = "N" And strCheck = "Y" Then
        'MICRO결과가 없을경우
        strCheck = "U"
    End If
    
    
    
    
    RES_MICROCHECK = strCheck
    

End Function

Function RES_CHECK(argEquipCode As String, argResult As String) As String
    Dim i As Integer
    Dim strLow As String
    Dim strHigh As String
    Dim strEqual As String
    RES_CHECK = "Y"
    
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT LOW, HIGH , EQUAL"
    SQL = SQL & vbCrLf & "  FROM EXAMMST"
    SQL = SQL & vbCrLf & " WHERE EQUIPCODE = '" & argEquipCode & "'"
    'SQL = SQL & vbCrLf & "   AND LOW <> ''"
    'SQL = SQL & vbCrLf & "ORDER BY INT(LOW)"
    res = db_select_Vas(gLocal, SQL, vasConvert)
    
    
    For i = 1 To vasConvert.DataRowCnt
        strLow = ""
        strHigh = ""
        strEqual = ""
        
        strLow = Trim(GetText(vasConvert, i, 1))
        strHigh = Trim(GetText(vasConvert, i, 2))
        strEqual = Trim(GetText(vasConvert, i, 3))
        
        If IsNumeric(strLow) = True And IsNumeric(strHigh) = True And Trim(strEqual) = "" Then
            If IsNumeric(argResult) = True Then
                If CCur(strLow) <= CCur(argResult) And CCur(strHigh) > CCur(argResult) Then
                   RES_CHECK = "N"
                   Exit For
                End If
            Else
                Exit For
            End If
        ElseIf IsNumeric(strLow) = True And IsNumeric(strHigh) = False And Trim(strEqual) = "" Then
            If IsNumeric(argResult) = True Then
                If CCur(strLow) <= CCur(argResult) Then
                    RES_CHECK = "N"
                    Exit For
                End If
            Else
                Exit For
            End If
        ElseIf IsNumeric(strLow) = False And IsNumeric(strHigh) = False And Trim(strEqual) <> "" Then
            If strEqual = argResult Then
                RES_CHECK = "N"
                Exit For
            End If
        End If
    Next i
    
    
End Function


Function RESULT_VALIDATE(argBarcode As String) As String
    RESULT_VALIDATE = "N"
    '결과 확정처리를 위한 함수
    '조건
    '1. KET 의 결과가 3+ 일경우에는 결과상태로 함.
    '2. MIRCRO만 있을경우 결과상태로 함.
    '3. 시험지오더만 있을있을때
    '  - KET이 3+ 이하이고, UBG, BIL 의 값이 NEGATIVE 일때 확정
    '4. 시험지+MICRO오더가 있을경우
    '  - ERY, LEU 값이 NEGATIVE 이고, KET이 3+ 이하이고, UBG, BIL 의 값이 NEGATIVE 일때 확정
    Dim i As Integer
    
    Dim strFinalState As String
    
    '각각의 오더값을 넣음
    Dim strKET As String
    Dim strUBG As String
    Dim strBIL As String
    Dim strNIT As String
    
    '각 변수의 값이 Y 일때는 최종보고 해도 된다는 표시
    Dim strKETFCheck As String
    Dim strUBGFCheck As String
    Dim strBILFCheck As String
    Dim strNITFCheck As String
    
    '각 변수의 값이 Y 일때는 처방이 있다는 표시
    Dim strKETOrd As String
    Dim strUBGOrd As String
    Dim strBILOrd As String
    Dim strNITOrd As String
    
    '각변수의 값이 있을때는 처방이 있다고 생각함.
    Dim strURINEORD As String
    Dim strMICROORD As String
    
    
    '최종보고
    strFinalState = "N"
    
On Error GoTo errCHECK
    
    
    '시험지 오더가 있는지...
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT TESTSUBCODE "
    SQL = SQL & vbCrLf & "  FROM LABREGRESULT "
    SQL = SQL & vbCrLf & " WHERE TESTSUBCODE IN (" & gAllExam_NAF & ")"
    SQL = SQL & vbCrLf & "   AND LABREGDATE = '" & "20" & Format(Mid(argBarcode, 1, 6), "@@-@@-@@") & "'"
    SQL = SQL & vbCrLf & "   AND LABREGNO = '" & Mid(argBarcode, 7) & "'"
    SQL = SQL & vbCrLf & " GROUP BY TESTSUBCODE "
    res = db_select_Row(gServer, SQL)
    strURINEORD = ""
    For i = 0 To res
        If Trim(gReadBuf(i)) <> "" Then
            If strURINEORD = "" Then
                strURINEORD = "'" & gReadBuf(i) & "'"
            Else
                strURINEORD = strURINEORD & ",'" & gReadBuf(i) & "'"
            End If
        End If
    Next i
    
    'MICRO 오더가 있는지...
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT TESTSUBCODE "
    SQL = SQL & vbCrLf & "  FROM LABREGRESULT "
    SQL = SQL & vbCrLf & " WHERE TESTSUBCODE IN (" & gAllExam_Micro & ")"
    SQL = SQL & vbCrLf & "   AND LABREGDATE = '" & "20" & Format(Mid(argBarcode, 1, 6), "@@-@@-@@") & "'"
    SQL = SQL & vbCrLf & "   AND LABREGNO = '" & Mid(argBarcode, 7) & "'"
    SQL = SQL & vbCrLf & " GROUP BY TESTSUBCODE "
    res = db_select_Row(gServer, SQL)
    strMICROORD = ""
    For i = 0 To res
        If Trim(gReadBuf(i)) <> "" Then
            If strMICROORD = "" Then
                strMICROORD = "'" & gReadBuf(i) & "'"
            Else
                strMICROORD = strMICROORD & ",'" & gReadBuf(i) & "'"
            End If
        End If
    Next i
    
    
    '결과값 체크
    'KET
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT RESVALUE, DELTAFLAG "
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'KET'"
    SQL = SQL & vbCrLf & "   AND BARCODE  = '" & argBarcode & "'"
    SQL = SQL & vbCrLf & " GROUP BY RESVALUE, DELTAFLAG "
    res = db_select_Col(gLocal, SQL)
    strKET = gReadBuf(0)
    strKETOrd = gReadBuf(1)
    strKETFCheck = "N"
    
    'UBG
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT RESVALUE, DELTAFLAG "
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'UBG'"
    SQL = SQL & vbCrLf & "   AND BARCODE  = '" & argBarcode & "'"
    SQL = SQL & vbCrLf & " GROUP BY RESVALUE, DELTAFLAG "
    res = db_select_Col(gLocal, SQL)
    strUBG = gReadBuf(0)
    strUBGOrd = gReadBuf(1)
    strUBGFCheck = "N"
    
    'BIL
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT RESVALUE, DELTAFLAG "
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'BIL'"
    SQL = SQL & vbCrLf & "   AND BARCODE  = '" & argBarcode & "'"
    SQL = SQL & vbCrLf & " GROUP BY RESVALUE, DELTAFLAG "
    res = db_select_Col(gLocal, SQL)
    strBIL = gReadBuf(0)
    strBILOrd = gReadBuf(1)
    strBILFCheck = "N"
    
    'NIT
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT RESVALUE, DELTAFLAG "
    SQL = SQL & vbCrLf & "  FROM PAT_RES"
    SQL = SQL & vbCrLf & " WHERE EQUIPCODE = 'NIT'"
    SQL = SQL & vbCrLf & "   AND BARCODE  = '" & argBarcode & "'"
    SQL = SQL & vbCrLf & " GROUP BY RESVALUE, DELTAFLAG "
    res = db_select_Col(gLocal, SQL)
    strNIT = gReadBuf(0)
    strNITOrd = gReadBuf(1)
    strNITFCheck = "N"
    
    'KET, UBG, BIL 결과 및 오더 유무를 확인하여 조건에 만족하는지 확인함.
    If strKETOrd = "" Then
        If strKET = "3+" Then
            strKETFCheck = "N"
        Else
            strKETFCheck = "Y"
        End If
    Else
        strKETFCheck = "Y"
    End If
    
    If strUBGOrd = "" Then
        If strUBG = "Negative" Then
            strUBGFCheck = "Y"
        Else
            strUBGFCheck = "N"
        End If
    Else
        strUBGFCheck = "Y"
    End If
    
    If strBILOrd = "" Then
        If strBIL = "Negative" Then
            strBILFCheck = "Y"
        Else
            strBILFCheck = "N"
        End If
    Else
        strBILFCheck = "Y"
    End If
    
    
    If strNITOrd = "" Then
        If strNIT = "Negative" Then
            strNITFCheck = "Y"
        Else
            strNITFCheck = "N"
        End If
    Else
        strNITFCheck = "Y"
    End If
    
    'URINE 오더가 없고 MICRO 오더만 있는 경우 결과 입력상태로 한다.
    If strMICROORD <> "" And strURINEORD = "" Then
        ' strFinalState = "N" 기본값.
        
        RESULT_VALIDATE = strFinalState
        Exit Function
    ElseIf strMICROORD = "" And strURINEORD <> "" Then
        If strKETFCheck = "Y" And strUBGFCheck = "Y" And strBILFCheck = "Y" Then
            RESULT_VALIDATE = "Y"
        Else
            RESULT_VALIDATE = "N"
        End If
    Else
        'URINE 오더가 있고 MICRO 오더가 있거나 없는경우 아래 조건을 확인후에 확정을 할지 말지 결정한다.
        'KET, UBG, BIL 체크값이 Y 일 경우에만 확정을 해준다.
        If strKETFCheck = "Y" And strUBGFCheck = "Y" And strBILFCheck = "Y" And strNITFCheck = "Y" Then
            SQL = ""
            SQL = SQL & vbCrLf & "SELECT RESVALUE"
            SQL = SQL & vbCrLf & "  FROM PAT_RES"
            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(argBarcode) & "'"
            SQL = SQL & vbCrLf & "   AND RESVALUE IN ('1+','2+','3+','4+','5+','Trace') "
            SQL = SQL & vbCrLf & "   AND EQUIPCODE = 'ERY' "
            SQL = SQL & vbCrLf & "UNION ALL"
            SQL = SQL & vbCrLf & "SELECT RESVALUE"
            SQL = SQL & vbCrLf & "  FROM PAT_RES"
            SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(argBarcode) & "'"
            SQL = SQL & vbCrLf & "   AND RESVALUE IN ('1+','2+','3+','4+','5+','Trace') "
            SQL = SQL & vbCrLf & "   AND EQUIPCODE = 'LEU' "
            res = db_select_Row(gLocal, SQL)
            
            'POSITIVE 결과가 있다면 N 으로 한다.
            If res > 0 Then
                RESULT_VALIDATE = "N"
            Else
                RESULT_VALIDATE = "Y"
            End If
            Exit Function
        Else
            ' strFinalState = "N" 기본값.
            RESULT_VALIDATE = strFinalState
            Exit Function
        End If
    End If
    
    Exit Function
    
errCHECK:
    '에러가 날경우 입력상태로 한다.
    RESULT_VALIDATE = "N"
    
End Function


Function CONVERT_RESULT(argEquipCode As String, argResult As String) As String
    CONVERT_RESULT = ""
    '특정 검사항목의 결과를 설정한 값으로 변경 시켜줌
    Dim i As Integer
    Dim strLow As String
    Dim strHigh As String
    Dim strEqual As String
    Dim strChageVal As String
    
    
    
    ClearSpread vasConvert
    
    strChageVal = ""
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT LOW, HIGH , EQUAL, VALSTRING"
    SQL = SQL & vbCrLf & "  FROM EXAMMST"
    SQL = SQL & vbCrLf & " WHERE EQUIPCODE = '" & argEquipCode & "'"
    'SQL = SQL & vbCrLf & "   AND LOW <> ''"
    'SQL = SQL & vbCrLf & "ORDER BY INT(LOW)"
    res = db_select_Vas(gLocal, SQL, vasConvert)
    
    If res < 1 Then
        CONVERT_RESULT = argResult
        Exit Function
    Else
        'If IsNumeric(Trim(Replace(Replace(argResult, ">", ""), "<", ""))) = False Then
            'CONVERT_RESULT = argResult
            'Exit Function
        'Else
            argResult = Trim(Replace(Replace(argResult, ">", ""), "<", ""))
        'End If
        
        'SG의 값이 - 가 나오는 경우 1.000 으로 바꿔줌
        If argEquipCode = "SG" And argResult = "-" Then
            argResult = "-"
        End If
    End If
        
    
    For i = 1 To vasConvert.DataRowCnt
        strLow = ""
        strHigh = ""
        strChageVal = ""
        
        strLow = Trim(GetText(vasConvert, i, 1))
        strHigh = Trim(GetText(vasConvert, i, 2))
        strEqual = Trim(GetText(vasConvert, i, 3))
        strChageVal = Trim(GetText(vasConvert, i, 4))
        
        If IsNumeric(strLow) = True And IsNumeric(strHigh) = True And Trim(strEqual) = "" Then
            If IsNumeric(argResult) = True Then
                If CCur(strLow) <= CCur(argResult) And CCur(strHigh) > CCur(argResult) Then
                   argResult = strChageVal
                   Exit For
                End If
            Else
                Exit For
            End If
        ElseIf IsNumeric(strLow) = True And IsNumeric(strHigh) = False And Trim(strEqual) = "" Then
            If IsNumeric(argResult) = True Then
                If CCur(strLow) <= CCur(argResult) Then
                    argResult = strChageVal
                    Exit For
                End If
            Else
                Exit For
            End If
        ElseIf IsNumeric(strLow) = False And IsNumeric(strHigh) = False And Trim(strEqual) <> "" Then
            If strEqual = argResult Then
                argResult = strChageVal
                Exit For
            End If
        End If
    Next i
    
    CONVERT_RESULT = argResult
    
End Function

Private Function Result_FLAG(asResult As String, asEQUIPCODE As String) As String
    Result_FLAG = ""
    
    Select Case asEQUIPCODE
        Case "WBC"
            If asResult <> "0 - 1" Then
                Result_FLAG = "*"
            End If
        Case "RBC"
            If asResult <> "0 - 1" Then
                Result_FLAG = "*"
            End If
        Case "EPCEL"
            If asResult <> "0 - 1" Then
                Result_FLAG = "*"
            End If
        Case "BAC"
            If asResult <> "Not found" Then
                Result_FLAG = "*"
            End If
        Case "CRY"
            If asResult <> "Not found" Then
                Result_FLAG = "*"
            End If
        Case "PAT"
            If asResult <> "Not found" Then
                Result_FLAG = "*"
            End If
        Case "OTR"
            If asResult <> "Not found" Then
                Result_FLAG = "*"
            End If
'/----------------------------------------------------------------------------
        Case "ERY"
            If asResult = "Negative" Then
            
            Else
                Result_FLAG = "*"
            End If
        Case "LEU"
            If asResult = "Negative" Then
            
            Else
                Result_FLAG = "*"
            End If
        Case "NIT"
            If asResult = "Negative" Then
            
            Else
                Result_FLAG = "*"
            End If
        Case "KET"
            If asResult = "Negative" Then
            
            Else
                Result_FLAG = "*"
            End If
        Case "GLU"
            If asResult = "Negative" Then
            
            Else
                Result_FLAG = "*"
            End If
        Case "PRO"
            If asResult = "Negative" Then
            
            Else
                Result_FLAG = "*"
            End If
        Case "UBG"
            If asResult = "Negative" Then
            
            Else
                Result_FLAG = "*"
            End If
        Case "BIL"
            If asResult = "Negative" Then
            
            Else
                Result_FLAG = "*"
            End If
        Case "pH"
            If IsNumeric(asResult) = True Then
                If CCur(asResult) > 8 Then
                    Result_FLAG = "H"
                ElseIf CCur(asResult) < 5 Then
                    Result_FLAG = "L"
                End If
            End If
        Case "SG"
            If IsNumeric(asResult) = True Then
                If CCur(asResult) > 1.035 Then
                    Result_FLAG = "H"
                ElseIf CCur(asResult) < 1.003 Then
                    Result_FLAG = "L"
                End If
            End If
    End Select
    
    
End Function

Private Function Result_Set(asExamCode As String, asResult As String, asEQUIPCODE As String) As String
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
    
    On Error GoTo ErrRes:
    
    Result_Set = ""
    
    strResValue = asResult
    
    If IsNumeric(strResValue) = False Then
        Result_Set = strResValue & "|"
        Exit Function
    End If
    
    For i = 1 To 11
        gReadBuf(i - 1) = ""
    Next
    
    SQL = "SELECT REPLOW, REPHIGH, REFLOW, REFHIGH, LSTRING, MSTRING, HSTRING, LEQUIL, HEQUIL, RESPREC, RESGUBUN " & vbCrLf & _
          "FROM EQUIPEXAM WHERE EQUIPCODE = '" & asEQUIPCODE & "' AND EXAMCODE = '" & asExamCode & "' "
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
    
    
    If IsNumeric(strRespRec) = True And strRespRec <> "9" And strResGubun <> "2" Then
        
        Select Case strRespRec
        
            Case "0"
                strResValue = Format(strResValue, "###0")
            Case "1"
                strResValue = Format(strResValue, "###0.0")
            Case "2"
                strResValue = Format(strResValue, "###0.00")
            Case "3"
                strResValue = Format(strResValue, "###0.000")
            Case "4"
                strResValue = Format(strResValue, "###0.0000")
            Case "5"
                strResValue = Format(strResValue, "###0.00000")
            Case Else
                strResValue = strResValue
        End Select
    Else
        strResValue = strResValue
    End If
    
    If IsNumeric(cRepL) = True Then
        If CCur(cRepL) >= CCur(strResValue) Then
            'strGiho = "<"
            'strResValue = cRepL
            
            
            If IsNumeric(strRespRec) = True And strRespRec <> "9" Then
                
                Select Case strRespRec
                
                    Case "0"
                        cRepL = Format(cRepL, "###0")
                    Case "1"
                        cRepL = Format(cRepL, "###0.0")
                    Case "2"
                        cRepL = Format(cRepL, "###0.00")
                    Case "3"
                        cRepL = Format(cRepL, "###0.000")
                    Case "4"
                        cRepL = Format(cRepL, "###0.0000")
                    Case "5"
                        cRepL = Format(cRepL, "###0.00000")
                    Case Else
                        cRepL = cRepL
                End Select
            Else
                cRepL = cRepL
            End If
            
            strResult = "<" & cRepL
            
            Result_Set = strResult & "|" & strResValue
            Exit Function
        End If
    End If
    
    If IsNumeric(cRepH) = True Then
        If CCur(cRepH) <= CCur(strResValue) Then
            'strGiho = ">"
            'strResValue = cRepH
            If IsNumeric(strRespRec) = True And strRespRec <> "9" Then
                
                Select Case strRespRec
                
                    Case "0"
                        cRepH = Format(cRepH, "###0")
                    Case "1"
                        cRepH = Format(cRepH, "###0.0")
                    Case "2"
                        cRepH = Format(cRepH, "###0.00")
                    Case "3"
                        cRepH = Format(cRepH, "###0.000")
                    Case "4"
                        cRepH = Format(cRepH, "###0.0000")
                    Case "5"
                        cRepH = Format(cRepH, "###0.00000")
                    Case Else
                        cRepH = cRepH
                End Select
            Else
                cRepH = cRepH
            End If
            
            strResult = ">" & cRepH
            Result_Set = strResult & "|" & strResValue
            Exit Function
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
    Else
        strResult = strResValue
        
    End If
    
    Result_Set = strResult & "|" & strGiho & strResValue
    Exit Function
    
ErrRes:
    
    Result_Set = strResValue & "|" & strGiho & strResValue
    Exit Function
    
End Function

Private Sub Init_Form()
    frmInterface.Caption = gEquipName & " Interface Program"
    SSPanel1.Caption = "     " & gEquipName & "  INTERFACE"
End Sub

Private Sub Command10_Click()
    SQL = "SELECT fn_sd_get_age('20160330','20160126') from dual"
    res = db_select_Col(gServer, SQL)
    MsgBox gReadBuf(0)
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

Private Sub Command4_Click()
        Winsock1.Close
    WinSock_Listen Winsock1
End Sub

Private Sub Command5_Click()
    Dim sPName As String
    
'    spname = "aaa"
''
'    spname = ConvertStringToUtf8Bytes(spname)
    'sPName = StrConv(Trim(GetText(vasID, glRow, colPName)), vbUnicode)
'    sPName = URLDecodeUTF8(Trim(sPName))
    
    Text1 = EncodeUTF8_ADOStream(chrVT & Text1 & chrFS & chrCR)


    Winsock1.SendData (EncodeUTF8_ADOStream(chrVT & Text1 & chrFS & chrCR))
'    Winsock1.SendData chrVT & Text1 & chrFS & chrCR
End Sub

Private Sub Command7_Click()
    'Cobas8000 (txtWinSockBuff)
    
    Dim sTmp As String
    Dim strSendData
    Dim strResFlag
    Dim sSigFlag As String
    Dim sStemp As String
    
    Dim arrSIGNAL
    Dim intX    As Long
    
On Error GoTo errCHECK
    
    
    For intX = 1 To Len(Text1)
        
        DoSleep 1
        sTmp = Mid(Text1, intX, 1)
        
        'Save_Raw_Data "[RX" & Format(time, "hh:mm:ss") & "]" & sTmp
        
       ' txtWinSockBuff = txtWinSockBuff & sTmp
        
        
        If InStr(1, sTmp, chrACK) = 0 Then
            txtWinSockBuff = txtWinSockBuff & sTmp
        End If
        
        If InStr(1, sTmp, chrENQ) > 0 Then
            blTimerChk = False
            Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & vbCrLf & txtWinSockBuff
            
            'Call Winsock1.SendData(chrACK)
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
            
            
            txtWinSockBuff = ""
            'Exit Sub
        End If
        
        If InStr(1, sTmp, chrLF) > 0 Then
'            If InStr(1, txtWinSockBuff, "[RX") > 0 Then
'                txtWinSockBuff = Mid(txtWinSockBuff, InStr(1, txtWinSockBuff, "[") + 12)
'                Save_Raw_Data "[RX" & Format(time, "hh:mm:ss") & "]" & vbCrLf & txtWinSockBuff
                
                'Call Winsock1.SendData(chrACK)
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
                
                
                
                
                If InStr(1, sTmp, "2Q") > 0 Then
                    strQMode = "Q"
                End If
            'End If
            'Exit Sub
        End If
        
        If InStr(1, sTmp, chrACK) > 0 Then
            'Call Winsock1.SendData(GetText(vasOrder, 1, 1))
            
            If gOrderMsg(0) = "" And gOrderMsg(1) = "" And gOrderMsg(2) = "" And gOrderMsg(3) = "" And gOrderMsg(4) <> "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(4)
                blTimerChk = True
                blTimerTerminate = True
                'Call Winsock1.SendData(gOrderMsg(4))
                gOrderMsg(4) = ""
                
                'If GetText(vasOrderTest, 2, 1) = "" Then
                DeleteRow vasOrderTest, 1, 1
                'End If
                
            ElseIf gOrderMsg(0) = "" And gOrderMsg(1) = "" And gOrderMsg(2) = "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(3)
                'Call Winsock1.SendData(gOrderMsg(3))
                gOrderMsg(3) = ""
                
            ElseIf gOrderMsg(0) = "" And gOrderMsg(1) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(2)
                'Call Winsock1.SendData(gOrderMsg(2))
                gOrderMsg(2) = ""
                
            ElseIf gOrderMsg(0) = "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(1)
                'Call Winsock1.SendData(gOrderMsg(1))
                gOrderMsg(1) = ""
               
            ElseIf gOrderMsg(0) <> "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(0)
                'Call Winsock1.SendData(gOrderMsg(0))
                gOrderMsg(0) = ""
               
                
            '끝나는 신호가 없을경우
            ElseIf gOrderMsg(0) <> "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(0)
                'Call Winsock1.SendData(gOrderMsg(0))
                gOrderMsg(0) = ""
            ElseIf gOrderMsg(0) = "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(1)
                'Call Winsock1.SendData(gOrderMsg(1))
                gOrderMsg(1) = ""
            ElseIf gOrderMsg(0) = "" And gOrderMsg(1) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(2)
                'Call Winsock1.SendData(gOrderMsg(2))
                gOrderMsg(2) = ""
                DeleteRow vasOrderTest, 1, 1
                FN_OrderSend
'''            Else
'''                Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(0)
'''                Call Winsock1.SendData(gOrderMsg(0))
'''                gOrderMsg(0) = ""
            End If
            
            'DeleteRow vasOrder, 1, 1
            
            
            'Exit Sub
        End If
        
        If InStr(1, sTmp, chrEOT) > 0 Then
            
            Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & txtWinSockBuff
            
            sSigFlag = Cobas8000(txtWinSockBuff)
            txtWinSockBuff = ""
            
    '
            If strQMode = "Q" Then
                'Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrENQ
                'Call Winsock1.SendData(chrENQ)
                
                strQMode = ""
            End If
            
            If vasOrderTest.DataRowCnt > 0 Then
                intTimer = 0
            End If
            
'''            If vasOrder.DataRowCnt > 0 Then
'''                intTimer = 0
'''            End If
            
    '''        If vasOrder.DataRowCnt > 0 Then
    '''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrENQ
    '''            Call Winsock1.SendData(chrENQ)
    '''        End If
            blTimerChk = True
            
            'Exit Sub
        End If
        
        
''        If InStr(1, sTmp, chrENQ) > 0 Then
''            'Call Winsock1.SendData(chrACK)
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrACK
''            'Exit Sub
''        End If
''
''        If InStr(1, sTmp, chrLF) > 0 Then
''            'Call Winsock1.SendData(chrACK)
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrACK
''            'Exit Sub
''        End If
''
''        If InStr(1, sTmp, chrACK) > 0 Then
''            'Call Winsock1.SendData(GetText(vasOrder, 1, 1))
''
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & GetText(vasOrder, 1, 1)
''            DeleteRow vasOrder, 1, 1
''            'Exit Sub
''        End If
''
''        If InStr(1, sTmp, chrEOT) > 0 Then
''            sSigFlag = Cobas8000(txtWinSockBuff)
''            Save_Raw_Data "[RX" & Format(time, "hh:mm:ss") & "]" & vbCrLf & txtWinSockBuff
''
''            txtWinSockBuff = ""
''
''            If vasOrder.DataRowCnt > 0 Then
''                'Call Winsock1.SendData(chrENQ)
''            End If
''
''            'Exit Sub
''        End If
    Next intX
    
    Exit Sub
    
errCHECK:
    Save_Raw_Data "[ERR WinsoCK]" & Format(Time, "hh:mm:ss")
    
    
End Sub

Private Sub Command8_Click()
'선택전송
Dim vasIDRow As Integer
Dim VasResRow As Integer
Dim iRow As Integer
Dim liRet As Integer
Dim iNumber As Integer

    If MsgBox(" " & vbCrLf & "선택전송을 하시겠습니까?" & vbCrLf & " ", vbInformation + vbOKCancel, "알림:추가오더선택전송") = vbCancel Then
        Exit Sub
    End If
    
    If vasList.DataRowCnt < 1 Then
        MsgBox "저장할 데이터가 없습니다."
        Exit Sub
    End If
    
    'db_BeginTran gServer
'''    Connect_Server
    For vasIDRow = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = vasIDRow
        
        If vasList.Value = 1 Then '체크된 열은 저장이 안됨
'        If vasID.Value = "" Then
        
            liRet = -1
            liRet = Insert_Data_ADD(vasIDRow, vasList)
            '''liRet = ToServer_Re(vasIDRow)
'''            liRet = Insert_Data_SHINWON(vasIDRow, vasList, Format(dtpExamDate, "yyyymmdd"))
            
            If liRet = 1 Then
                'db_Commit gServer
                
                SetBackColor vasList, vasIDRow, vasIDRow, colBarCode, colState, 255, 255, 180
                SetText vasList, "Trans", vasIDRow, colState
            Else
                SetBackColor vasList, vasIDRow, vasIDRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasList, "Failed", vasIDRow, colState
            End If
            vasList.Col = 1
            vasList.Row = vasIDRow
            vasList.Value = 0
        Else
        
        End If
    Next vasIDRow
End Sub

Private Sub Command9_Click()
    'Cobas8000 (txtWinSockBuff)
    
    Dim sTmp As String
    Dim strSendData
    Dim strResFlag
    Dim sSigFlag As String
    Dim sStemp As String
    
    Dim arrSIGNAL
    Dim intX    As Long
    
On Error GoTo errCHECK
    
    
    For intX = 1 To Len(Text1)
        
        DoSleep 1
        sTmp = Mid(Text1, intX, 1)
        If InStr(1, sTmp, chrACK) = 0 Then
            txtWinSockBuff = txtWinSockBuff & sTmp
        End If
        
        If InStr(1, sTmp, chrENQ) > 0 Then
            blTimerChk = False
            
            Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & txtWinSockBuff
            txtWinSockBuff = ""
            
            'Call Winsock1.SendData(chrACK)
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
            'Exit Sub
        End If
        
        If InStr(1, sTmp, chrLF) > 0 Then
            
            'Call Winsock1.SendData(chrACK)
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
            
            If InStr(1, sTmp, "2Q") > 0 Then
                strQMode = "Q"
            End If
            'Exit Sub
        End If
        
        If InStr(1, sTmp, chrACK) > 0 Then
            
            If gOrderMsg(0) = "" And gOrderMsg(2) = "" And gOrderMsg(3) = "" And gOrderMsg(4) <> "" Then
                    Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(4)
                    blTimerChk = True
                    blTimerTerminate = True
                    'Call Winsock1.SendData(gOrderMsg(4))
                    gOrderMsg(4) = ""
                    
                    'If GetText(vasOrderTest, 2, 1) = "" Then
                        DeleteRow vasOrderTest, 1, 1
                    'End If
                    
            ElseIf gOrderMsg(0) = "" And gOrderMsg(2) = "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(3)
                'Call Winsock1.SendData(gOrderMsg(3))
                gOrderMsg(3) = ""
                
            ElseIf gOrderMsg(0) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(2)
                'Call Winsock1.SendData(gOrderMsg(2))
                gOrderMsg(2) = ""
                
    '        ElseIf gOrderMsg(0) = "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
    '            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(1)
    '            Call Winsock1.SendData(gOrderMsg(1))
    '            gOrderMsg(1) = ""
                
            ElseIf gOrderMsg(0) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(0)
                'Call Winsock1.SendData(gOrderMsg(0))
                gOrderMsg(0) = ""
                
            '끝나는 신호가 없을경우
            ElseIf gOrderMsg(0) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(0)
                'Call Winsock1.SendData(gOrderMsg(0))
                gOrderMsg(0) = ""
    '        ElseIf gOrderMsg(0) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
    '            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(1)
    '            Call Winsock1.SendData(gOrderMsg(1))
    '            gOrderMsg(1) = ""
            ElseIf gOrderMsg(0) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(2)
                'Call Winsock1.SendData(gOrderMsg(2))
                gOrderMsg(2) = ""
                DeleteRow vasOrderTest, 1, 1
                FN_OrderSend
            End If
            
            'Exit Sub
        End If
        
        If InStr(1, sTmp, chrEOT) > 0 Then
            Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & txtWinSockBuff
            
            sSigFlag = Cobas8000(txtWinSockBuff)
            txtWinSockBuff = ""
    
            If strQMode = "Q" Then
                'Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrENQ
                'Call Winsock1.SendData(chrENQ)
                
                strQMode = ""
            End If
            
            If vasOrderTest.DataRowCnt > 0 Then
                intTimer = 0
                
            End If
            blTimerChk = True

            'Exit Sub
        End If

    Next intX
    
    Exit Sub
    
errCHECK:
    Save_Raw_Data "[ERR WinsoCK]" & Format(Time, "hh:mm:ss")
    
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
    
    frmInterface.Width = 16425
    frmInterface.Height = 11520
    
    cmdClear_Click
    
    ClearSpread vasList
    
    GetSetup    'ini에서 DB정보 불러오기
    
    Init_Form
    intErrorCheck = 0
    '나중에 풀어야함.
''
    If Not Connect_Server Then
        MsgBox "연결되지 않았습니다."
        Exit Sub
    End If
    
        '\\127.0.0.1\in'
    
    'TMAX 연결 =========================================
    'gTMAX.TUX_READENV gDB_Parm.EnvPath, gDB_Parm.User
'''    gTMAX.TP_INIT
    '===================================================

    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        Exit Sub
    End If
    
    'FN_Compact
    '수정후 꼭 풀기
    
    WinSock_Listen Winsock1
    
    ClearSpread vasOrder
    
    blTimerChk = True
    intTimer = 0
    
'    lblUser.Caption =
'    txtUID.text = gExamUID

    raw_data = ""
    gExamUID = "POCT"
    txtUID.Text = gExamUID
    
    dtpToday = Now
    dtpExamDate = dtpToday
    dtpExamDate2 = dtpToday
    dtpAddExamDate1 = dtpToday
    dtpAddExamDate2 = dtpToday
    '====================로컬 DB지우기 - 30일 보관======================
    sDate = Format(DateAdd("y", CDate(dtpToday), -gLocalExpDate), "yyyymmdd")
    
    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    res = SendQuery(gLocal, SQL)
    '===================================================================

    '검사코드 가져오기
    GetExamCode

    ClearSpread vasCode
    
    vasID.MaxRows = 1
    vasID.ColsFrozen = 6
    vasRes.MaxRows = 20
    vasList.MaxRows = 1
    
    vasList.ColsFrozen = 6
    
    vasLISTRES.MaxRows = 20
    
    dtpSDate = Date - 2
    dtpEDate = Date
    
'''    vasID.Visible = False
    For i = colRStart + 1 To vasID.MaxCols
        vasID.Col = i
        vasID.ColHidden = True
    Next
    
    For i = 1 To vasID.MaxCols
        vasID.Row = 0
        vasID.Col = i
        vasID.FontSize = 9
    Next
    
'''    vasID.Visible = True
    
'''    vasList.Visible = False
    For i = colRStart + 1 To vasList.MaxCols
        vasList.Col = i
        vasList.ColHidden = True
    Next
'''    vasID.Visible = True

    
    SSTab1.Tab = 0
    
    cmbTransGubun.ListIndex = 1
    
    SQL = "update equipexam set equipno = '" & gEquip & "' "
    res = SendQuery(gLocal, SQL)
    
    cmdVasIDWidth_Click
    cmdVasIDWidth.Visible = False
    
    gRecordCnt = 1
    gPatCnt = 1
    strHeader = "H|\^&|||c65002^Cobas6500^2.0.0^7^SU0500151^SV0500082|||||||P|LIS2-A2|" & chrCR & chrETX
    strTerminate = "L|1|N" & chrCR & chrETX
    gOrderMsg(0) = ""
    gOrderMsg(1) = ""
    gOrderMsg(2) = ""
    gOrderMsg(3) = ""
    gOrderMsg(4) = ""
    
    cboMicro.AddItem "전체", 0
    cboMicro.AddItem "Y", 1
    cboMicro.AddItem "U", 2
    cboMicro.AddItem "N", 3
    cboMicro.ListIndex = 0
    lblInfo1.Caption = ""
    lblInfo2.Caption = ""
    'strTerminate = chrSTX & strTerminate & CheckSum(strTerminate) & chrCR & chrLF
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    'Unload frmLogin
    Unload frmExamSearch
    WritePrivateProfileString "config", "UID", txtUID.Text, App.Path & "\interface.ini"
'''    gTMAX.TP_TERM
'''    DisConnect_Server
    KillProcess "IF_Cobas.exe"
    DisConnect_Local
    End
End Sub

Sub GetExamCode()
'검사코드를 array에 저장
    Dim i As Integer
    
    gAllExam = ""
    gOrderExam = ""
    gReceExam = ""
    gAllExam_NAF = ""
    
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

    SQL = "Select ExamCode From EquipExam  "

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
    
    ClearSpread vasTemp

    SQL = "Select ExamCode From EquipExam WHERE deltavalue = '1' "

    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    For i = 1 To vasTemp.DataRowCnt

        If Trim(GetText(vasTemp, i, 1)) <> "" Then
            If gAllExam_NAF = "" Then
                gAllExam_NAF = "'" & Trim(GetText(vasTemp, i, 1)) & "'"
            Else
                gAllExam_NAF = gAllExam_NAF & ",'" & Trim(GetText(vasTemp, i, 1)) & "'"
            End If
        End If
    Next i
    
    
    ClearSpread vasTemp

    SQL = "Select ExamCode From EquipExam WHERE deltavalue = '0' "

    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    For i = 1 To vasTemp.DataRowCnt

        If Trim(GetText(vasTemp, i, 1)) <> "" Then
            If gAllExam_Micro = "" Then
                gAllExam_Micro = "'" & Trim(GetText(vasTemp, i, 1)) & "'"
            Else
                gAllExam_Micro = gAllExam_Micro & ",'" & Trim(GetText(vasTemp, i, 1)) & "'"
            End If
        End If
    Next i
    
    If gAllExam_Micro = "" Then
        gAllExam_Micro = "''"
    End If
    
    
End Sub

Private Sub Label1_Click()
    If Frame5.Visible = True Then
        Frame5.Visible = False
        vasListTemp.Visible = False
        vasOrderTest.Visible = False
    Else
        Frame5.Visible = True
        vasListTemp.Visible = True
        vasOrderTest.Visible = True
    End If
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

Private Sub mnuExamMst_Click()
    frmExamMst.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload frmErrorList
    Unload Me
End Sub

Private Sub mnuManual_Click()
    mnuManual.Checked = True
    mnuAuto.Checked = False
End Sub

'''Private Sub MSComm1_OnComm()
'''
'''    Dim s As String
'''    Dim sPID As String
'''    Dim sSendData As String
'''    Dim sSndMessage As String
'''    Dim i As Integer
'''    Dim iRow As Integer
'''    Dim lResRow As Long
'''
'''    Dim sExamCode As String
'''    Dim sExamName As String
'''    Dim sResult As String
'''
'''    s = MSComm1.Input
'''
'''    Select Case s
'''
'''    Case chrENQ
'''        Save_Raw_Data "[Rx" & Format(time, "hh:mm:ss") & "]" & chrENQ
'''
'''        gSndState = ""
'''        gENQFlag = 9
'''
'''        gRecodeType = ""
'''        MSComm1.Output = chrACK
'''        Save_Raw_Data "[Tx" & Format(time, "hh:mm:ss") & "]" & chrACK
'''
'''        gPreSpecID = ""
'''        gPreRow = 0
'''
'''    Case chrACK
'''
'''        Save_Raw_Data "[Rx" & Format(time, "hh:mm:ss") & "]" & chrACK
'''        gOrdRow = gOrdRow + 1
'''
'''        If GetText(vasOrder, gOrdRow, 1) = "" Then
'''            Exit Sub
'''        End If
'''
'''        If gOrdRow <= vasOrder.DataRowCnt Then
'''
'''            sSendData = Trim(GetText(vasOrder, gOrdRow, 1))
'''
'''            MSComm1.Output = sSendData
'''            Save_Raw_Data "[Tx" & Format(time, "hh:mm:ss") & "]" & sSendData
'''
'''            If gOrdRow = vasOrder.DataRowCnt Then
'''                ClearSpread vasOrderBuf
'''                ClearSpread vasOrder
'''
'''                Me.MousePointer = 0
'''            End If
'''        End If
'''
'''    Case chrSTX     '자료수신 시작
'''        txtData.text = s
'''
'''    Case chrETX
'''        txtData.text = txtData.text & s
'''
'''    Case chrLF
'''        txtData.text = txtData.text & s
'''        Save_Raw_Data "[Rx" & Format(time, "hh:mm:ss") & "]" & txtData.text
'''
'''        Architect txtData.text
'''
'''        MSComm1.Output = chrACK
'''        Save_Raw_Data "[Tx" & Format(time, "hh:mm:ss") & "]" & chrACK
'''
'''    Case chrEOT     '자료수신 완료
'''        If gRecodeType = "R" Then
'''
'''            gSndState = "R"
'''
'''        ElseIf gRecodeType = "Q" Then
'''            gOrdRow = 0
'''            gPreMsg = chrENQ
'''
'''            frmInterface.MSComm1.Output = chrENQ
'''            Save_Raw_Data "[Tx" & Format(time, "hh:mm:ss") & "]" & chrENQ
'''
'''            gSndState = "Q"
'''            gPreMsg = chrENQ
'''        End If
'''
'''        gMsgFlag = ""
'''        gHeadRecode = ""
'''        txtData.text = ""
'''
'''    Case Else
'''        txtData.text = txtData.text & s
'''    End Select
'''End Sub
'''Sub SendOrder()
'''Dim sSendOrder As String
''''''
''''''    gOrderCnt = 1
'''
'''    If Len(gOrderMessage) > 240 Then
'''
'''        If gOrderCnt = 8 Then
'''            gOrderCnt = 0
'''        End If
'''
'''        sSendOrder = CStr(gOrderCnt) & Left(gOrderMessage, 240) & chrETB
'''        gOrderMessage = Mid(gOrderMessage, 241)
'''
'''        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
'''        SaveQuery sSendOrder, 1
'''
'''        gOrderCnt = gOrderCnt + 1
'''        comSend = "stENQ"
'''
'''        gPreMsg = sSendOrder
'''        Save_Raw_Data "[TX]" & sSendOrder
'''
'''        MSComm1.Output = sSendOrder
'''
'''    Else
'''        If gOrderCnt = 8 Then
'''            gOrderCnt = 0
'''        End If
'''
'''        sSendOrder = CStr(gOrderCnt) & gOrderMessage & chrETX
'''        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
'''
'''        gOrderMessage = ""
'''        comSend = "stOrder"
'''
'''        gPreMsg = sSendOrder
'''        Save_Raw_Data "[TX]" & sSendOrder
'''
'''        MSComm1.Output = sSendOrder
'''    End If
'''End Sub

Function Make_Order(asSpecid As String, asRow As Long, asRackPos As String) As String

    Dim sDate As String
    
    Dim sCnt As String
    
    Dim sOCnt As Long
    Dim sRetOrder As String
    Dim sOrder As String
    
    Dim iRow As Long
    Dim llRow As Long
    Dim llRow_Order As Long
    
    Dim sBarCode As String
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
    Dim strRackPos As String
    Dim strArr() As String
    Dim sSType As String
    
    Dim strORDERType As String
    
    Dim intFinalR As Integer
    '/효준 추가
    Dim gstrOrder(5)
    Dim intX    As Integer
    
    Dim lsOrder As String
    
    Dim strDept As String
    
    Dim strALLOrder As String
    
    vasOrderBuf.MaxRows = 0
    lsOrder = ""
    
   
''    If Len(asSpecid) <> 11 Then Exit Function
    
    lsReceDate = "20" & Mid(asSpecid, 1, 2) & "/" & Mid(asSpecid, 3, 2) & "/" & Mid(asSpecid, 5, 2)
    lsReceNo = Trim(GetText(vasID, glRow, colWkNo))
    strDept = Trim(GetText(vasID, glRow, colDept))
    
        '나중에 풀어야함.
'''
'''    If Len(asSpecid) = 11 Then
'''
'''        If IsNumeric(lsReceNo) = True And IsDate(lsReceDate) Then
'''            lsReceNo = CStr(CCur(lsReceNo))
'''            SQL = "Interface_GetPatientResult '" & gWorkList & "', '" & lsReceDate & "', " & lsReceNo & " "
'''            res = db_select_VasS(gServer, SQL, vasListTemp)
'''
'''        End If
'''    End If
'''
'''    '최종결과처리가 1개 이상이고 Micro오더만 있을때에는 Urine 검사를 하지 않음
'''    intFinalR = 0
'''    lsAllReceCode = ""
'''    For i = 1 To vasListTemp.DataRowCnt
'''        If Trim(GetText(vasListTemp, i, 20)) <> "최종" Then
'''            If lsAllReceCode = "" Then
'''                lsAllReceCode = "'" & Trim(GetText(vasListTemp, i, 11)) & "'"
'''            Else
'''                lsAllReceCode = lsAllReceCode & ", '" & Trim(GetText(vasListTemp, i, 11)) & "'"
'''            End If
'''        ElseIf Trim(GetText(vasListTemp, i, 20)) = "최종" Then
'''            intFinalR = intFinalR + 1
'''        End If
'''    Next
    '전체오더확인
    ClearSpread vasListTemp
    If IsDate(Format(Mid(asSpecid, 1, 6), "@@-@@-@@")) = True Then
    
        SQL = "SELECT R.ORDCODE "
        SQL = SQL & vbCrLf & "   FROM SLRSLTMT R"
        SQL = SQL & vbCrLf & " WHERE 1 = 1"
        SQL = SQL & vbCrLf & "   AND R.SPCDATE = TO_DATE(SUBSTR('" & asSpecid & "', 1, 6), 'YYMMDD')"
        SQL = SQL & vbCrLf & "   AND R.SPCNO = SUBSTR('" & asSpecid & "', 7, 5)"
        SQL = SQL & vbCrLf & "   AND R.SPCSEQ = SUBSTR('" & asSpecid & "', 12, 1)"
        'AND A.PROCSTAT IN ('D', 'E')"
        'SQL = SQL & vbCrLf & "   AND R.PROCSTAT IN ('D', 'E')"
        'SQL = SQL & vbCrLf & "   AND R.ORDCODE IN (" & strExamList & ")"
        SQL = SQL & vbCrLf & " GROUP BY R.ORDCODE "
        res = db_select_Vas(gServer, SQL, vasListTemp)
    End If
    lsAllReceCode = ""
    For i = 1 To vasListTemp.DataRowCnt
        If lsAllReceCode = "" Then
            lsAllReceCode = "'" & Trim(GetText(vasListTemp, i, 1)) & "'"
        Else
            lsAllReceCode = lsAllReceCode & ",'" & Trim(GetText(vasListTemp, i, 1)) & "'"
        End If
    Next i
    
    If lsAllReceCode = "" Then
        lsAllReceCode = "''"
    End If
    
    ClearSpread vasTemp
    SQL = "select equipcode,'',  examcode, EXAMNAME,SEQNO from EQUIPEXAM "
    SQL = SQL & vbCrLf & "where  examcode in (" & lsAllReceCode & ") "
    SQL = SQL & vbCrLf & "GROUP BY equipcode, examcode, EXAMNAME,SEQNO "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    strALLOrder = ""
    If res = 0 Then
        'SetText vasID, "오더없음", glRow, colState
        'Exit Function
    Else
        SQL = "select equipcode,  examcode, EXAMNAME from EQUIPMST "
        SQL = SQL & vbCrLf & "where  examcode in (" & lsAllReceCode & ") "
        SQL = SQL & vbCrLf & "GROUP BY equipcode,  examcode, EXAMNAME "
        res = db_select_Row(gLocal, SQL)
        
        For i = 1 To res
            If InStr(1, strALLOrder, gReadBuf(i - 1)) = 0 Then
                If strALLOrder = "" Then
                    strALLOrder = gReadBuf(i - 1)
                Else
                    strALLOrder = strALLOrder & gReadBuf(i - 1)
                End If
            End If
        Next i
    End If
    
    
    
    '미완료오더확인.
    ClearSpread vasListTemp
    If IsDate(Format(Mid(asSpecid, 1, 6), "@@-@@-@@")) = True Then
    
        SQL = "SELECT R.ORDCODE "
        SQL = SQL & vbCrLf & "   FROM SLRSLTMT R"
        SQL = SQL & vbCrLf & " WHERE 1 = 1"
        SQL = SQL & vbCrLf & "   AND R.SPCDATE = TO_DATE(SUBSTR('" & asSpecid & "', 1, 6), 'YYMMDD')"
        SQL = SQL & vbCrLf & "   AND R.SPCNO = SUBSTR('" & asSpecid & "', 7, 5)"
        SQL = SQL & vbCrLf & "   AND R.SPCSEQ = SUBSTR('" & asSpecid & "', 12, 1)"
        'AND A.PROCSTAT IN ('D', 'E')"
        SQL = SQL & vbCrLf & "   AND R.PROCSTAT IN ('D', 'E')"
        'SQL = SQL & vbCrLf & "   AND R.ORDCODE IN (" & strExamList & ")"
        SQL = SQL & vbCrLf & " GROUP BY R.ORDCODE "
        res = db_select_Vas(gServer, SQL, vasListTemp)
    End If
    lsAllReceCode = ""
    For i = 1 To vasListTemp.DataRowCnt
        If lsAllReceCode = "" Then
            lsAllReceCode = "'" & Trim(GetText(vasListTemp, i, 1)) & "'"
        Else
            lsAllReceCode = lsAllReceCode & ",'" & Trim(GetText(vasListTemp, i, 1)) & "'"
        End If
    Next i
    
    If lsAllReceCode = "" Then
        lsAllReceCode = "''"
    End If
    
    ClearSpread vasTemp
    SQL = "select equipcode,'',  examcode, EXAMNAME,SEQNO from EQUIPEXAM "
    SQL = SQL & vbCrLf & "where  examcode in (" & lsAllReceCode & ") "
    SQL = SQL & vbCrLf & "GROUP BY equipcode, examcode, EXAMNAME,SEQNO "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    strORDERType = ""
    If res = 0 Then
        'SetText vasID, "오더없음", glRow, colState
        'Exit Function
    Else
        SQL = "select equipcode,  examcode, EXAMNAME from EQUIPMST "
        SQL = SQL & vbCrLf & "where  examcode in (" & lsAllReceCode & ") "
        SQL = SQL & vbCrLf & "GROUP BY equipcode,  examcode, EXAMNAME "
        res = db_select_Row(gLocal, SQL)
        
        For i = 1 To res
            If InStr(1, strORDERType, gReadBuf(i - 1)) = 0 Then
                If strORDERType = "" Then
                    strORDERType = gReadBuf(i - 1)
                Else
                    strORDERType = strORDERType & gReadBuf(i - 1)
                End If
            End If
        Next i
    End If
    
    
    
    
    '전체오더중 미완료 오더가 포함되어 있을때 전체오더에 따르게 한다.
    If strORDERType <> "" Then
        If InStr(1, strALLOrder, strORDERType) > 0 Then
            strORDERType = strALLOrder
        End If
    End If
    
    
    '검사를 진행 할때
    'Urine만 있을경우는 Urine만 진행 --> C
    'Urine, Micro 가있을경우 Sieve 검사를 진행 --> S
    'Micro 가있을경우 Urine,Micro 진행 --> CM
    If InStr(1, strORDERType, "C") > 0 And InStr(1, strORDERType, "M") = 0 Then
        strORDERType = "C"
        If Trim(GetText(vasID, glRow, colState)) = "" Then
            SetBackColor vasID, glRow, glRow, colOrdStick, colOrdStick, 255, 50, 50
            SetText vasID, "1", glRow, colStickCheck
        End If
    ElseIf InStr(1, strORDERType, "C") > 0 And InStr(1, strORDERType, "M") > 0 Then
        
        '당분간은 CM으로 검사를 한다.
        'strOrderType = "S"
        strORDERType = "CM"
        If Trim(GetText(vasID, glRow, colState)) = "" Then
            SetBackColor vasID, glRow, glRow, colOrdStick, colOrdStick, 255, 50, 50
            SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 255, 50, 50
            SetText vasID, "1", glRow, colStickCheck
            SetText vasID, "1", glRow, colMicroCheck
        End If
    ElseIf InStr(1, strORDERType, "C") = 0 And InStr(1, strORDERType, "M") > 0 Then


        strORDERType = "M"
        If Trim(GetText(vasID, glRow, colState)) = "" Then
            SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 255, 50, 50
            SetText vasID, "1", glRow, colStickCheck
        End If

    ElseIf InStr(1, strORDERType, "C") = 0 And InStr(1, strORDERType, "M") = 0 Then
        strORDERType = ""
        If strALLOrder <> "" And strORDERType = "" Then
            SetText vasID, "완료검체", glRow, colState
        Else
            If Trim(GetText(vasID, glRow, colState)) = "" Then
                If IsNumeric(asSpecid) = True And Len(asSpecid) = 12 Then
                    SetText vasID, "오더없음", glRow, colState
                Else
                    SetText vasID, "Bar.Err", glRow, colState
                    SetBackColor vasID, glRow, glRow, colRackPos, colState, 200, 50, 255
                End If
            End If
        End If
        
    End If
    
    
    
    
    
    
'''
'''    If InStr(1, strOrderType, "C") > 0 And InStr(1, strOrderType, "M") = 0 Then
'''        strOrderType = "C"
'''        If Trim(GetText(vasID, glRow, colState)) = "" Then
'''            SetBackColor vasID, glRow, glRow, colOrdStick, colOrdStick, 255, 50, 50
'''            SetText vasID, "1", glRow, colStickCheck
'''        End If
'''
'''
'''    ElseIf InStr(1, strOrderType, "C") > 0 And InStr(1, strOrderType, "M") > 0 Then
'''        strOrderType = "S"
'''        If Trim(GetText(vasID, glRow, colState)) = "" Then
'''            SetBackColor vasID, glRow, glRow, colOrdStick, colOrdStick, 255, 50, 50
'''            SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 255, 50, 50
'''            SetText vasID, "1", glRow, colStickCheck
'''            SetText vasID, "1", glRow, colMicroCheck
'''        End If
'''
'''
'''    ElseIf InStr(1, strOrderType, "C") = 0 And InStr(1, strOrderType, "M") > 0 Then
'''
'''        'If intFinalR > 0 Then
'''            strOrderType = "M"
'''            If Trim(GetText(vasID, glRow, colState)) = "" Then
'''                SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 255, 50, 50
'''                SetText vasID, "1", glRow, colStickCheck
'''            End If
'''        'Else
'''        '    strOrderType = "CM"
'''        '    If Trim(GetText(vasID, glRow, colState)) = "" Then
'''        '        SetBackColor vasID, glRow, glRow, colOrdStick, colOrdStick, 255, 50, 50
'''        '        SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 255, 50, 50
'''        '        SetText vasID, "1", glRow, colStickCheck
'''        '        SetText vasID, "1", glRow, colMicroCheck
'''        '    End If
'''        'End If
'''    Else
'''
'''    End If
    '나중에 풀어야함.
''        SQL = ""
''        SQL = SQL & vbCrLf & "SELECT COUNT(TESTCODE)"
''        SQL = SQL & vbCrLf & "  FROM LabRegResult  "
''        SQL = SQL & vbCrLf & " WHERE LabRegDate = '" & lsReceDate & "'"
''        SQL = SQL & vbCrLf & "   AND TestCode in (SELECT TestCode FROM labtestcode WHERE SampleCode = '03' AND TestCode NOT LIKE ('17%'))"
''        SQL = SQL & vbCrLf & "   AND labregno = '" & lsReceNo & "'"
''        res = db_select_Col(gServer, SQL)
'''
'''        If Len(asSpecid) = 11 Then
'''
'''            If gReadBuf(0) <> "" Then
'''                If IsNumeric(gReadBuf(0)) = True And CCur(gReadBuf(0)) = 0 Then
'''                    If intFinalR > 0 Then
'''                        strOrderType = ""
'''                        SetText vasID, "최종", glRow, colState
'''                        SetBackColor vasID, glRow, glRow, colOrdStick, colOrdStick, 255, 255, 255
'''                        SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 255, 255, 255
'''                    Else
'''                        strOrderType = "S"
'''                        SetText vasID, "미접수", glRow, colState
'''                        SetBackColor vasID, glRow, glRow, colOrdStick, colOrdStick, 200, 50, 255
'''                        SetBackColor vasID, glRow, glRow, colOrdMicro, colOrdMicro, 200, 50, 255
'''                    End If
'''                Else
'''                    strOrderType = ""
'''                    If Trim(GetText(vasID, glRow, colState)) = "" Then
'''                        SetText vasID, "오더없음", glRow, colState
'''                    End If
'''                End If
'''            End If
'''        Else
'''            strOrderType = ""
'''            If Trim(GetText(vasID, glRow, colState)) = "" Then
'''                SetText vasID, "오더없음", glRow, colState
'''            End If
'''        End If
'''
'''    End If
    
    If Trim(GetText(vasID, glRow, colPName)) = "" And Len(asSpecid) = 11 Then
        Get_Sample_Info glRow
    End If
    
    sPID = Trim(GetText(vasID, glRow, colPID))
    sPName = Trim(GetText(vasID, glRow, colPName))
    strRackPos = Trim(GetText(vasID, glRow, colRackPos))
    sSex = Trim(GetText(vasID, glRow, colPSex))
    
    If sAge = "" Then
        sAge = "0"
    End If
    
    iCnt = 0
    'ClearSpread vasOrder
    ClearSpread vasOrderBuf

    'Order 생성하기==================================================

    sEmgFlag = lsFlag
    
    llRow_Order = 1

    gCurMsgCnt = 1
    
    'HeadH|\^&|||cobas-e411^1|||||host|RSUPL^REAL|P|1
    'gHeader = "H|\^&|||c65002^Cobas6500^2.0.0^7^SU0500151^SV0500082|||||||P|LIS2-A2|"
    'gHeader = "H|\^&|||host^2|||||H7600|TSDWN^REPLY|P|1"
    'gPatient = "P|1"
    
    sReqEquipCode = ""
    For i = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, i, 1)) <> "" Then
                If InStr(1, sReqEquipCode, Trim(GetText(vasTemp, i, 1))) = 0 Then
                    sReqEquipCode = sReqEquipCode & Trim(GetText(vasTemp, i, 1))
                End If
        End If
        
        sEquipCode = Trim(GetText(vasTemp, i, 1))
        sEquipNo = Trim(GetText(vasTemp, i, 2))
        sExamCode = Trim(GetText(vasTemp, i, 3))
        sExamName = Trim(GetText(vasTemp, i, 4))
        sSeqNo = Trim(GetText(vasTemp, i, 5))
'''        SQL = "SELECT EQUIPCODE, EXAMNAME, SEQNO FROM EQUIPEXAM " & vbCrLf & _
'''              "WHERE equipno = '" & sEquipNo & "' and  EXAMCODE = '" & sExamCode & "' "
'''        res = db_select_Col(gLocal, SQL)
        
'''        sEquipCode = Trim(gReadBuf(0))

        
        SetPositionResult asRow, sEquipCode, "*"
        
        SQL = "select barcode from pat_res where barcode = '" & Trim(asSpecid) & "' and examcode = '" & Trim(sExamCode) & "' AND EQUIPCODE = '" & sEquipCode & "'"
        SQL = SQL & vbCrLf & " AND EXAMDATE = '" & Format(Date, "yyyymmdd") & "'"
        res = db_select_Col(gLocal, SQL)
        If res = 0 Then
            
            SQL = ""
            SQL = SQL & vbCrLf & "insert into pat_res(equipno, examdate, "
            SQL = SQL & vbCrLf & "                    barcode, examcode, equipcode,"
            SQL = SQL & vbCrLf & "                    result, resvalue, pname, pid, "
            SQL = SQL & vbCrLf & "                    seqno, page, examname, posno, "
            SQL = SQL & vbCrLf & "                    psex, sendflag, receno, EXAMUID, DISKNO) "
            SQL = SQL & vbCrLf & "values('" & gEquip & "', '" & Format(Date, "yyyymmdd") & "', "
            SQL = SQL & vbCrLf & "       '" & Trim(asSpecid) & "','" & Trim(sExamCode) & "','" & Trim(sEquipCode) & "',"
            SQL = SQL & vbCrLf & "       '', '', '" & Trim(sPName) & "', '" & Trim(sPID) & "',"
            SQL = SQL & vbCrLf & "       '" & Trim(sSeqNo) & "', '" & Trim(sAge) & "', '" & Trim(sExamName) & "', '" & Trim(strRackPos) & "',"
            SQL = SQL & vbCrLf & "       '" & Trim(sSex) & "','0','" & lsReceNo & "','" & strDept & "','" & strORDERType & "')"
            res = SendQuery(gLocal, SQL)
        End If
            
    Next
    
    If Trim(GetText(vasID, glRow, colState)) = "오더없음" Then
        SQL = ""
        SQL = SQL & vbCrLf & "insert into pat_res(equipno, examdate, "
        SQL = SQL & vbCrLf & "                    barcode, examcode, equipcode,"
        SQL = SQL & vbCrLf & "                    result, resvalue, pname, pid, "
        SQL = SQL & vbCrLf & "                    seqno, page, examname, posno, "
        SQL = SQL & vbCrLf & "                    psex, sendflag, receno, EXAMUID, DISKNO) "
        SQL = SQL & vbCrLf & "values('" & gEquip & "', '" & Format(Date, "yyyymmdd") & "', "
        SQL = SQL & vbCrLf & "       '" & Trim(asSpecid) & "','오더없음','오더없음',"
        SQL = SQL & vbCrLf & "       '', '', '" & Trim(sPName) & "', '" & Trim(sPID) & "',"
        SQL = SQL & vbCrLf & "       '" & Trim(sSeqNo) & "', '" & Trim(sAge) & "', '" & "" & "', '" & Trim(strRackPos) & "',"
        SQL = SQL & vbCrLf & "       '" & Trim(sSex) & "','6','" & lsReceNo & "','" & strDept & "','" & strORDERType & "')"
        res = SendQuery(gLocal, SQL)
    End If
    
    lsOrder = strORDERType
    
    '/Head         1H|\^&||||||||||P||(CR)59
    gstrOrder(0) = "1H|\^&|||c65002^Cobas6500^2.0.0^7^SU0500151^SV0500082|||||||P|LIS2-A2|" & chrCR & chrETX
    gstrOrder(0) = chrSTX & gstrOrder(0) & CheckSum(gstrOrder(0)) & chrCR & chrLF
    
'''    'P부분이 없음
'''    '/Patient       2P|1|||||||||||||||||||||||||||||||||(CR)96
'''    gstrOrder(1) = "2P|1" & chrCR & chrETX
'''    gstrOrder(1) = chrSTX & gstrOrder(1) & CheckSum(gstrOrder(1)) & chrCR & chrLF
                                                                                            '|R||||||N||||||||||||||O
    '/Order        3O|1|00000031    |^4^3                                 |^^^410^0\^^^900^0|R||||||N||||||||||||||O(CR)44
    '                O|1|000663      |36^0044^2^^SAMPLE^NORMAL|ALL|R|20050705093416|||||X||||||||||||||O
    '               3O|1|" & gSPID & "|" & gSampleNo & "^" & gRack & "^" & gPos & "|" & sOrder & "|R||||||N||||||||||||||O" & chrCR & chrETX
    'gstrOrder(2) = "3O|1|" & lsID & "|" & txtStartNo & "^0^" & txtPosNo & "|"
    
    
    If Trim(GetText(vasID, glRow, colState)) = "미접수" Or Trim(GetText(vasID, glRow, colState)) = "오더없음" Then
        If strORDERType = "" Then
                        'O|1|14102901931|500197^5^User^SAMPLE|CM|R||||||N|||20120508115956|||||||||||Q
            gstrOrder(1) = "2O|1|" & asSpecid & "|" & asRackPos & "^User^SAMPLE|"
            gstrOrder(1) = gstrOrder(1) & lsOrder & "|R||||||N||||||||||||||Q" & chrCR & chrETX
            gstrOrder(1) = chrSTX & gstrOrder(1) & CheckSum(gstrOrder(1)) & chrCR & chrLF
            
            If Trim(GetText(vasID, glRow, colState)) = "" Then
                SetText vasID, "오더없음", glRow, colState
            End If
            
        Else
            gstrOrder(1) = "2O|1|" & asSpecid & "|" & asRackPos & "|"
            gstrOrder(1) = gstrOrder(1) & lsOrder & "|R||||||N||||||||||||||Q" & chrCR & chrETX
            gstrOrder(1) = chrSTX & gstrOrder(1) & CheckSum(gstrOrder(1)) & chrCR & chrLF
            
            If Trim(GetText(vasID, glRow, colState)) = "" Or Trim(GetText(vasID, glRow, colState)) = "Trans" Then
                SetText vasID, "오더", glRow, colState
            End If
            
        End If
        
    Else
    
        If vasTemp.DataRowCnt = 0 Then
                        'O|1|14102901931|500197^5^User^SAMPLE|CM|R||||||N|||20120508115956|||||||||||Q
            gstrOrder(1) = "2O|1|" & asSpecid & "|" & asRackPos & "^User^SAMPLE|"
            gstrOrder(1) = gstrOrder(1) & "|R||||||N||||||||||||||Q" & chrCR & chrETX
            gstrOrder(1) = chrSTX & gstrOrder(1) & CheckSum(gstrOrder(1)) & chrCR & chrLF
            
            If Trim(GetText(vasID, glRow, colState)) = "" Then
                SetText vasID, "오더없음", glRow, colState
            End If
            
        Else
            gstrOrder(1) = "2O|1|" & asSpecid & "|" & asRackPos & "|"
            gstrOrder(1) = gstrOrder(1) & lsOrder & "|R||||||N||||||||||||||Q" & chrCR & chrETX
            gstrOrder(1) = chrSTX & gstrOrder(1) & CheckSum(gstrOrder(1)) & chrCR & chrLF
            
            If Trim(GetText(vasID, glRow, colState)) = "" Or Trim(GetText(vasID, glRow, colState)) = "Trans" Then
                SetText vasID, "오더", glRow, colState
            End If
            
        End If
    End If
    '/terminater    4L|1|N(CR)3D
    gstrOrder(2) = "3L|1|N" & chrCR & chrETX
    gstrOrder(2) = chrSTX & gstrOrder(2) & CheckSum(gstrOrder(2)) & chrCR & chrLF
    
    '/EOT           
    gstrOrder(3) = chrEOT
    
    
    If vasOrder.DataRowCnt > 0 Then
        vasOrder.MaxRows = vasOrder.DataRowCnt + 1
        SetText vasOrder, chrENQ, vasOrder.DataRowCnt + 1, 1
    End If
    
    'If vasOrderTest.DataRowCnt = 0 Then
        vasOrderTest.MaxRows = vasOrderTest.DataRowCnt + 1
        
        SetText vasOrderTest, asSpecid, vasOrderTest.MaxRows, 1
        SetText vasOrderTest, lsOrder, vasOrderTest.MaxRows, 2
        SetText vasOrderTest, asRackPos, vasOrderTest.MaxRows, 3
    'End If
    
    
    
    
'''    '/오더 넣기(SPREAD)
'''    vasOrder.MaxRows = vasOrder.DataRowCnt + 1
'''    SetText vasOrder, gstrOrder(0), vasOrder.DataRowCnt + 1, 1
'''    vasOrder.MaxRows = vasOrder.DataRowCnt + 1
'''    SetText vasOrder, gstrOrder(1), vasOrder.DataRowCnt + 1, 1
'''    vasOrder.MaxRows = vasOrder.DataRowCnt + 1
'''    SetText vasOrder, gstrOrder(2), vasOrder.DataRowCnt + 1, 1
'''    vasOrder.MaxRows = vasOrder.DataRowCnt + 1
'''    SetText vasOrder, gstrOrder(3), vasOrder.DataRowCnt + 1, 1
    
    'vasOrder.MaxRows = vasOrder.DataRowCnt + 1
    'SetText vasOrder, gstrOrder(4), vasOrder.DataRowCnt + 1, 1

    
End Function

Private Sub SetPositionResult(asRow As Long, asEQUIPCODE As String, asResult As String)
    Dim strEquipCode As String
    Dim strResult As String
    Dim lngRow As Long
    Dim i As Integer
    
    lngRow = asRow
    strEquipCode = asEQUIPCODE
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
    Dim lsBarcode As String
    Dim lRow As Long
    Dim lsQCGubun As String
    
    Get_Sample_Info = -1
    lsQCGubun = False
    
    '샘플 환자 정보 가져오기
    lsBarcode = Mid(Trim(GetText(vasID, asRow, colBarCode)), 1, 11)   '샘플 바코드 번호
    
    'If IsNumeric(lsBarcode) = False Or Len(lsBarcode) < 10 Then Exit Function
    If Mid(Trim(lsBarcode), 1, 1) = "9" Then
        lsQCGubun = True
    End If
    
    
    lRow = asRow
    
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
    Dim lsReceDate As String
    Dim strErr As String
    
    '환자정보 가져오기
    sID = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
    If IsDate(Format(Mid(sID, 1, 6), "@@-@@-@@")) = False Then Exit Function
    
    SQL = "    SELECT DISTINCT P.PATNO, P.PATNAME,  TO_CHAR(R.WORKDATE, 'YYYYMMDD') || '-' ||R.LABNO"
    SQL = SQL & vbCrLf & "    ,(SELECT MEDDEPT || '-' ||WARDNO FROM SLACPTMT WHERE ROWNUM = 1 AND SPCDATE = R.SPCDATE AND SPCNO = R.SPCNO AND SPCSEQ = R.SPCSEQ )"
    SQL = SQL & vbCrLf & "    ,P.SEX"
    SQL = SQL & vbCrLf & "  FROM SLRSLTMT R, ACPATBAT P"
    SQL = SQL & vbCrLf & " WHERE R.SPCDATE = TO_DATE(SUBSTR('" & sID & "', 1, 6), 'YYMMDD')"
    SQL = SQL & vbCrLf & "   AND R.SPCNO = SUBSTR('" & sID & "', 7, 5)"
    SQL = SQL & vbCrLf & "   AND R.SPCSEQ = SUBSTR('" & sID & "', 12, 1)"
    SQL = SQL & vbCrLf & "   AND R.PATNO = P.PATNO"

    res = db_select_Col(gServer, SQL)
    If res = 1 Then

        'SetText vasID, Trim(gReadBuf(0)), asRow, colReceDate
        SetText vasID, Trim(gReadBuf(2)), glRow, colWkNo
        SetText vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText vasID, Trim(gReadBuf(1)), asRow, colPName
        SetText vasID, Trim(gReadBuf(4)), glRow, colPSex
        SetText vasID, Trim(gReadBuf(3)), asRow, colDept
        
    ElseIf res = 0 Then
    '환자정보가 없다면 접수처리를 한뒤 다시 조회한다.
        
        SQL = "SELECT SPCDATE, SPCNO, SPCSEQ "
        SQL = SQL & vbCrLf & "  From SLACPTMT "
        SQL = SQL & vbCrLf & " WHERE SPCDATE = TO_DATE('" & Mid(sID, 1, 6) & "', 'YYMMDD') "
        SQL = SQL & vbCrLf & "   AND SPCNO = '" & Mid(sID, 7, 5) & "'"
        SQL = SQL & vbCrLf & "   AND SPCSEQ = '" & Mid(sID, 12, 1) & "'"
        res = db_select_Col(gServer, SQL)
        
        
        If Len(sID) = 12 And IsDate(Format(Mid(sID, 1, 6), "@@-@@-@@")) = True Then
        strErr = spACPT_IDU(Trim(gReadBuf(0)), Trim(gReadBuf(1)), Trim(gReadBuf(2)), "POCT", Winsock1.LocalIP, Format(Now, "YYYY-MM-DD HH:MM:SS"))
            If Trim(strErr) = "N" Then
                DoSleep 500
                
                SQL = "    SELECT DISTINCT P.PATNO, P.PATNAME,  TO_CHAR(R.WORKDATE, 'YYYYMMDD') || '-' ||R.LABNO"
                SQL = SQL & vbCrLf & "    ,(SELECT MEDDEPT || '-' ||WARDNO FROM SLACPTMT WHERE ROWNUM = 1 AND SPCDATE = R.SPCDATE AND SPCNO = R.SPCNO AND SPCSEQ = R.SPCSEQ )"
                SQL = SQL & vbCrLf & "    ,P.SEX"
                SQL = SQL & vbCrLf & "  FROM SLRSLTMT R, ACPATBAT P"
                SQL = SQL & vbCrLf & " WHERE R.SPCDATE = TO_DATE(SUBSTR('" & sID & "', 1, 6), 'YYMMDD')"
                SQL = SQL & vbCrLf & "   AND R.SPCNO = SUBSTR('" & sID & "', 7, 5)"
                SQL = SQL & vbCrLf & "   AND R.SPCSEQ = SUBSTR('" & sID & "', 12, 1)"
                SQL = SQL & vbCrLf & "   AND R.PATNO = P.PATNO"
            
                res = db_select_Col(gServer, SQL)
                If res = 1 Then
            
                    'SetText vasID, Trim(gReadBuf(0)), asRow, colReceDate
                    SetText vasID, Trim(gReadBuf(2)), glRow, colWkNo
                    SetText vasID, Trim(gReadBuf(0)), asRow, colPID
                    SetText vasID, Trim(gReadBuf(1)), asRow, colPName
                    SetText vasID, Trim(gReadBuf(4)), glRow, colPSex
                    SetText vasID, Trim(gReadBuf(3)), asRow, colDept
                End If
            End If
        End If
    End If

    

End Function

'Function Get_Sample_Info(ByVal asRow As Long) As Integer
'    Dim sID As String
'
'    Dim lsPID As String
'    Dim lsPname As String
'    Dim lsDate As String
'    Dim lsSDate As String
'    Dim lsEDate As String
'    Dim iRow As Integer
'    Dim iCol As Integer
'    Dim strTmaxRes As String
'    Dim i As Long
'    Dim j As Long
'    Dim lsRow As Long
'    Dim strMsg As String
'    Dim sSmyr As String
'    Dim sSmsn As String
'    Dim sSms1 As String
'
'
'    '환자정보 가져오기
'    sID = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
'    lsDate = "20" & Mid(sID, 1, 6)
'    lsSDate = Format(dtpSDate, "yyyymmdd")
'    lsEDate = Format(dtpEDate, "yyyymmdd")
'
'    If sID = "" Then
'        Exit Function
'    End If
'
'    '바코드, 병록번호, 환자명, 검체코드, 검체명
'
'    ClearSpread vasTMaxList
''    select ACPTDD --일자
''        , ACPTNO -- 검체번호
''        , NM --성명
''        , RGSTNO --주민번호
''        , SEX --성별
''        , AGE --나이
''        , (select NM from shinbase..custcd c where a.CUSTCD_CD = c.CD) --거래처
''     from shinrslt..acptinfo a
''    where acptdd = '20121201' --일자
''      and acptno = '' --검체번호
'    SQL = ""
'    SQL = SQL & vbCrLf & "SELECT ACPTDD, ACPTNO, NM, '', SEX, AGE, "
'    SQL = SQL & vbCrLf & "      (select NM from shinbase..custcd c where a.CUSTCD_CD = c.CD)"
'    SQL = SQL & vbCrLf & "  from shinrslt..acptinfo a"
'    SQL = SQL & vbCrLf & " where acptdd = '" & lsDate & "' "
'    SQL = SQL & vbCrLf & "   and acptno = '" & Mid(sID, 7, 7) & "'"
''''    SQL = "SELECT A.SPCM_NO, A.PID , B.PT_NM , A.SPCM_CD , c.SPCM_ENM " & vbCrLf & _
''''          "FROM MS.MSLRCPT A " & vbCrLf & _
''''          "INNER JOIN MS.MSLGNRLRSLT AA ON A.RCPN_SQNO = AA.RCPN_SQNO " & vbCrLf & _
''''          "INNER JOIN HO.PCPPATIENT B ON A.PID = B.PID " & vbCrLf & _
''''          "INNER JOIN MS.MSLSPCMM C ON A.SPCM_CD = C.SPCM_CD " & vbCrLf & _
''''          "WHERE A.SPCM_NO = '" & sID & "' " & vbCrLf & _
''''          "AND AA.EXMN_CD IN (" & gAllExam & ") " & vbCrLf & _
''''          "GROUP BY A.SPCM_NO, A.PID, B.PT_NM, A.SPCM_CD, C.SPCM_ENM"
'    res = db_select_Col(gServer, SQL)
'
'    If res > 0 Then
'
'        SetText vasID, Trim(gReadBuf(1)), asRow, colPID
'        SetText vasID, Trim(gReadBuf(2)), asRow, colPName
'
'        If Trim(gReadBuf(4)) = "1" Then
'            SetText vasID, "M", asRow, colPsex
'        Else
'            SetText vasID, "F", asRow, colPsex
'        End If
'
'    End If
'
'End Function

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

Private Sub mnuOrder_Click()
    frmEquipMst.Show 1
End Sub

Private Sub mnuTransSet_Click()
    frmTransSet.Show 1
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


Private Sub subUp_Click()
    Dim sValue As String
    Dim sTmp As String
    Dim i As Integer
    Dim j As Integer
    
        sTmp = ""
    
        vasID.Row = vasID.ActiveRow
        vasID.Col = colBarCode
    
        sTmp = vasID.Text
        If sTmp = "" Then Exit Sub
        sValue = InputBox("변경할 검체번호를 입력하세요")
    
        If Trim(sValue) <> "" Then
            If MsgBox("" & sTmp & "를 " & sValue & "로 수정하시겠습니까?", vbYesNo, "확인") = vbYes Then
                SetText vasID, sValue, vasID.Row, vasID.Col
    
                If Trim(GetText(vasID, vasID.Row, colBarCode)) <> "" Then
                    Get_Sample_Info vasID.Row
    
                    For i = 1 To vasRes.DataRowCnt
                        Save_Local_One vasID.Row, i, "A"
                    Next
                End If
            End If
        End If
End Sub

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

Private Sub Timer1_Timer()
    If dtpToday <> Date Then
        dtpToday = Date
        ClearSpread vasID
        ClearSpread vasRes
    End If
    
End Sub

Private Sub Timer2_Timer()
    Dim ECHO As ICMP_ECHO_REPLY
    Dim pos As Integer
    
    Dim i As Integer
    
    Dim strTemp1 As String
    Dim strTemp2 As String
    
On Error GoTo ErrorCheck
    
    
    Call Ping(gEquipIP, ECHO)
   
    strTemp2 = GetStatusCode(ECHO.status)
    'strTemp2 = strTemp2 & vbCrLf & ECHO.Address
    'strTemp2 = strTemp2 & vbCrLf & ECHO.RoundTripTime & " ms"
    'strTemp2 = strTemp2 & vbCrLf & ECHO.DataSize & " bytes"
   
'    If Left$(ECHO.Data, 1) <> Chr$(0) Then
'       pos = InStr(ECHO.Data, Chr$(0))
'      strTemp2 = strTemp2 & vbCrLf & Left$(ECHO.Data, pos - 1)
'    End If
'
'    strTemp1 = ECHO.DataPointer

    If Trim(Mid(strTemp2, 1, InStr(1, strTemp2, "[") - 1)) = 0 Then
        Label6.Visible = False
    Else
        Label6.Visible = True
        Label6.Caption = "접속에러 : " & Winsock1.State
        Winsock1.Close
        WinSock_Listen Winsock1
        
    End If
    
'''    If intErrorCheck > 60 Then
'''        intErrorCheck = 0
'''        ClearSpread frmErrorList.vasErrorList
'''
'''
'''
'''        frmErrorList.Caption = gEquipName & " 미전송 결과 리스트"
'''
'''        SQL = ""
'''        SQL = SQL & vbCrLf & "DELETE FROM EXAMCHECK"
'''        SQL = SQL & vbCrLf & " WHERE EXAMDATE <= '" & Format(Date - 2, "YYYYMMDD") & "'"
'''        res = SendQuery(gLocal, SQL)
'''
'''
'''        SQL = ""
'''        SQL = SQL & vbCrLf & "SELECT A.BARCODE, RIGHT(A.BARCODE, 5), B.POSNO, '','', A.URINE, A.MICRO"
'''        SQL = SQL & vbCrLf & "  FROM EXAMCHECK A , PAT_RES B"
'''        SQL = SQL & vbCrLf & " WHERE A.BARCODE = B.BARCODE"
'''        'SQL = SQL & vbCrLf & "   AND B.EXAMDATE = MAX(B.EXAMDATE)"
'''        'SQL = SQL & vbCrLf & "   AND B.EXAMDATE = (SELECT EXAMDATE FROM PAT_RES WHERE BARCODE = A.BARCODE GROUP BY EXAMDATE )"
'''        SQL = SQL & vbCrLf & " GROUP BY A.BARCODE, RIGHT(A.BARCODE, 5), B.POSNO, A.URINE, A.MICRO"
'''        res = db_select_Vas(gLocal, SQL, frmErrorList.vasErrorList)
'''
'''
'''        If res > 0 Then
'''            frmErrorList.Left = 1000
'''            frmErrorList.Top = 1000
'''
'''
'''
'''            For i = 1 To frmErrorList.vasErrorList.DataRowCnt
'''
'''                If GetText(frmErrorList.vasErrorList, i, 6) = "Y" Then
'''                    SetBackColor frmErrorList.vasErrorList, i, i, 4, 4, 255, 50, 50
'''                End If
'''
'''                If GetText(frmErrorList.vasErrorList, i, 7) = "Y" Then
'''                    SetBackColor frmErrorList.vasErrorList, i, i, 5, 5, 255, 50, 50
'''                End If
'''
'''            Next i
'''
'''
'''            If frmErrorList.Visible <> True Then
'''                frmErrorList.Visible = True
'''                frmErrorList.Left = frmInterface.Left + 2000
'''            End If
'''        Else
'''            If frmErrorList.Visible <> False Then
'''                frmErrorList.Visible = False
'''            End If
'''        End If
        
        
'''    Else
'''        intErrorCheck = intErrorCheck + 1
'''    End If

Exit Sub

ErrorCheck:
Save_Raw_Data "Timer2 Error"
    
''    If Winsock1.State = 7 Then
''        Label6.Visible = False
''    Else
''        Label6.Visible = True
''        Label6.Caption = "접속에러 : " & Winsock1.State
''        Winsock1.Close
''        WinSock_Listen Winsock1
''
''    End If
End Sub

Private Sub Timer3_Timer()

    Dim steTemp As String
    
On Error GoTo errCHECK:
    If blTimerChk = False Then Exit Sub
    
    intTimer = intTimer + 1
    
    If intTimer > 3 Then
        intTimer = 0
    Else
        Exit Sub
    End If
    
    If Label6.Visible = True Or Winsock1.State <> 7 Then Exit Sub
'    steTemp = ""
'    Winsock1.GetData steTemp
'    If steTemp = chrENQ Then MsgBox "1"
    
    If vasOrderTest.DataRowCnt > 0 Then
        
        If blTimerTerminate = False And (gOrderMsg(0) <> "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "") Then
            gOrderMsg(0) = ""
            gOrderMsg(1) = ""
            gOrderMsg(2) = ""
            gOrderMsg(3) = ""
            gOrderMsg(4) = ""
            
            gRecordCnt = 1
            gPatCnt = 1
            FN_OrderSend
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrENQ
            Call Winsock1.SendData(chrENQ)
            blTimerChk = False
            blTimerTerminate = False
        Else
            gRecordCnt = 1
            gPatCnt = 1
            FN_OrderSend
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrENQ
            Call Winsock1.SendData(chrENQ)
            
            blTimerChk = False
            blTimerTerminate = False
        End If
    End If
    
''    If vasOrder.DataRowCnt > 0 Then
''        Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrENQ
''        Call Winsock1.SendData(chrENQ)
''        blTimerChk = False
''    End If
Exit Sub

errCHECK:
Save_Raw_Data "Error Timer3"
End Sub

Function FN_OrderSend() As String
    FN_OrderSend = "$$$$"
    
    If gRecordCnt = 8 Then
        gRecordCnt = 0
    End If
    
    If gPatCnt = 6 Then
        gPatCnt = 1
        gRecordCnt = 1
        Exit Function
    End If
    
    If gRecordCnt = 1 And gPatCnt = 1 Then
        gOrderMsg(0) = gRecordCnt & strHeader
        gOrderMsg(0) = chrSTX & gOrderMsg(0) & CheckSum(gOrderMsg(0)) & chrCR & chrLF
        gRecordCnt = gRecordCnt + 1
    Else
        gOrderMsg(0) = ""
    End If
    
'    gOrderMsg(1) = gRecordCnt & "P|" & gPatCnt & chrCR & chrETX
'    gOrderMsg(1) = chrSTX & gOrderMsg(1) & CheckSum(gOrderMsg(1)) & chrCR & chrLF
'    gRecordCnt = gRecordCnt + 1
    
'''    gOrderMsg(2) = gRecordCnt & "O|1|" & Trim(GetText(vasOrderTest, 1, 1)) & "|" & Trim(GetText(vasOrderTest, 1, 3)) & "^User^SAMPLE|"
'''
'''    If Trim(GetText(vasOrderTest, 1, 2)) = "" Then
'''
'''        gOrderMsg(2) = gOrderMsg(2) & Trim(GetText(vasOrderTest, 1, 2)) & "||||||||||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Y" & chrCR & chrETX
'''    Else
'''        gOrderMsg(2) = gOrderMsg(2) & Trim(GetText(vasOrderTest, 1, 2)) & "|R||||||N|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
'''    End If

    gOrderMsg(2) = gRecordCnt & "O|1|" & Trim(GetText(vasOrderTest, 1, 1)) & "|" & Trim(GetText(vasOrderTest, 1, 3)) & "^User^SAMPLE|"
    
    If Trim(GetText(vasOrderTest, 1, 2)) = "" Then
        If Len(Trim(GetText(vasOrderTest, 1, 1))) <> 12 Then
            If optOrder(0).Value = True Then 'Urine
                gOrderMsg(2) = gOrderMsg(2) & "C|R||||||N|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
            ElseIf optOrder(1).Value = True Then 'Sieve
                gOrderMsg(2) = gOrderMsg(2) & "S|R||||||N|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
            ElseIf optOrder(2).Value = True Then 'U + M
                gOrderMsg(2) = gOrderMsg(2) & "CM|R||||||N|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
                
            Else 'Cancel
                gOrderMsg(2) = gOrderMsg(2) & "|R||||||C|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
            End If
            
        Else
            'If optOrder(0).Value = True Then 'Urine
            '    gOrderMsg(2) = gOrderMsg(2) & "C|R||||||N|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
            'ElseIf optOrder(1).Value = True Then 'Sieve
            '    gOrderMsg(2) = gOrderMsg(2) & "S|R||||||N|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
            'ElseIf optOrder(2).Value = True Then 'U + M
            '    gOrderMsg(2) = gOrderMsg(2) & "CM|R||||||N|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
            '
            'Else 'Cancel
                gOrderMsg(2) = gOrderMsg(2) & "|R||||||C|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
            'End If
            
            ''바코드가 있는 경우
            'gOrderMsg(2) = gOrderMsg(2) & "S|R||||||N|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
            ''gOrderMsg(2) = gOrderMsg(2) & "||||||||||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Y" & chrCR & chrETX
        End If
        
    Else
        gOrderMsg(2) = gOrderMsg(2) & Trim(GetText(vasOrderTest, 1, 2)) & "|R||||||N|||" & Format(Now, "yyyymmddhhmmss") & "|||||||||||Q" & chrCR & chrETX
    End If
    
    gOrderMsg(2) = chrSTX & gOrderMsg(2) & CheckSum(gOrderMsg(2)) & chrCR & chrLF
    gRecordCnt = gRecordCnt + 1
    
    If Trim(GetText(vasOrderTest, 2, 1)) = "" Or gPatCnt = 5 Then
        gOrderMsg(3) = gRecordCnt & strTerminate
        gOrderMsg(3) = chrSTX & gOrderMsg(3) & CheckSum(gOrderMsg(3)) & chrCR & chrLF
        gRecordCnt = 1
        gPatCnt = 1
        gOrderMsg(4) = chrEOT
        Exit Function
    End If
    
    gPatCnt = gPatCnt + 1
End Function






Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = vbKeyReturn Then
        ClearSpread vasList
        
    
        SQL = "select '', FORMAT(examdate + MAX(EXAMTIME), '@@@@-@@-@@ @@:@@:@@'),  posno, MID(TRIM(receno),3), barcode,  pid,  pname, PSEX, TRIM(EXAMUID), '', '',EXAMTYPE from pat_res " & vbCrLf & _
              " where 1=1 "
        SQL = SQL & vbCrLf & "  and barcode like ('%" & Trim(txtBarcode.Text) & "%')"
        SQL = SQL & vbCrLf & " group by  barcode, posno, pid,MID(TRIM(receno),3), pname, PSEX , TRIM(EXAMUID),EXAMTYPE, examdate"
        res = db_select_Vas(gLocal, SQL, vasList)
    
        
        vasList.MaxRows = vasList.DataRowCnt
        For i = 1 To vasList.DataRowCnt
            SetVasColor vasList, CInt(i), Trim(GetText(vasList, i, colBarCode))
            If GetText(vasList, i, colState) = "1" Then
                SetText vasList, "Result", i, colState
                
            ElseIf GetText(vasList, i, colState) = "2" Then
                SetText vasList, "Trans", i, colState
                SetBackColor vasList, i, i, colBarCode, colState, 255, 255, 180
            End If
        Next
        
    End If
End Sub

'''Private Sub Timer1_Timer()
''''''    If dtpToday <> Date Then
''''''        dtpToday = Date
''''''    End If
'''
'''End Sub

Private Sub txtUID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gExamUID = txtUID.Text
        Call WritePrivateProfileString("CONFIG", "UID", txtUID.Text, App.Path & "\Interface.ini")
    End If
End Sub

Private Sub txtWorkNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
    
    If KeyCode = vbKeyReturn Then
        ClearSpread vasList
        
    
        SQL = "select '',FORMAT(examdate + MAX(EXAMTIME), '@@@@-@@-@@ @@:@@:@@'),  posno, MID(TRIM(receno),3), barcode,  pid,  pname, PSEX,  TRIM(EXAMUID), '', '',EXAMTYPE from pat_res " & vbCrLf & _
              " where 1=1 "
        SQL = SQL & vbCrLf & "  and receno like ('%" & Trim(txtWorkNo.Text) & "%')"
        SQL = SQL & vbCrLf & " group by  barcode, posno, pid,MID(TRIM(receno),3), pname, PSEX , TRIM(EXAMUID),EXAMTYPE, examdate"
        res = db_select_Vas(gLocal, SQL, vasList)
    
        
        vasList.MaxRows = vasList.DataRowCnt
        For i = 1 To vasList.DataRowCnt
            SetVasColor vasList, CInt(i), Trim(GetText(vasList, i, colBarCode))
            If GetText(vasList, i, colState) = "1" Then
                SetText vasList, "Result", i, colState
                
            ElseIf GetText(vasList, i, colState) = "2" Then
                SetText vasList, "Trans", i, colState
                SetBackColor vasList, i, i, colBarCode, colState, 255, 255, 180
            End If
        Next
        
    End If
End Sub

Private Sub vasBARERR_DblClick(ByVal Col As Long, ByVal Row As Long)
    vasBARERR.DeleteRows Row, 1
    
    If IsNumeric(Right(cmdBARERR.Caption, 2)) = True Then
        cmdBARERR.Caption = "BARCODE ERR - " & (Val(Right(cmdBARERR.Caption, 2)) - 1)
    Else
        cmdBARERR.Caption = "BARCODE ERR"
    End If
End Sub

'Private Sub Timer1_Timer()
'    Dim lRow As Long
'    Dim lCnt As Long
'    Dim sID As String
'    Dim sCode As String
'    Dim sDate As String
'    Dim sRack As String
'    Dim sTube As String
'    Dim sNew As String
'    Dim i As Long
'    Dim X As Integer
'
'    If ComState = False Then
'        Exit Sub
'    End If
'
''    Save_Raw_Data "[OrderCnt]" & vasCode.DataRowCnt
'    For i = 1 To vasCode.DataRowCnt
'        sID = Trim(GetText(vasCode, i, 3))
'        sCode = Trim(GetText(vasCode, i, 2))
'        sDate = Trim(GetText(vasCode, i, 4))
'        sRack = Trim(GetText(vasCode, i, 5))
'        sTube = Trim(GetText(vasCode, i, 6))
'        sNew = Trim(GetText(vasCode, i, 7))
'        If sCode <> "" And sID <> "" Then
'            Save_Raw_Data "[TimerCnt]" & vasCode.DataRowCnt
'            Integra800_Order_Entry sID, sDate, sCode, sRack, sTube, sNew
'            DeleteRow vasCode, i, i
'
'            Exit Sub
'        Else
'            DeleteRow vasCode, i, i
'            i = i - 1
'        End If
'    Next i
'
'    If Host_BC = "09" Then
'        For lRow = 1 To vasID.DataRowCnt
'            If InStr(1, Trim(GetText(vasID, lRow, 6)), "수신완료") > 0 Then
'                lCnt = lCnt + 1
'            Else
'            End If
'        Next lRow
'        If lCnt < vasID.DataRowCnt Then
'            Integra800_Res_Req
'            Integra800_QCRes_Req
'        Else
'            Integra800_OrderID_Req
'        End If
''            Integra800_QCRes_Req
'    ElseIf Left(Host_BC, 2) = "60" Or Host_BC = "00" Then
'        For lRow = 1 To vasID.DataRowCnt
'            If InStr(1, Trim(GetText(vasID, lRow, 6)), "수신완료") > 0 Then
'            'If InStr(1, Trim(GetText(vasID, lRow, colState)), "수신완료") > 0 Then
'                lCnt = lCnt + 1
'            Else
'            End If
'        Next lRow
'        If lCnt < vasID.DataRowCnt Then
'            Integra800_Res_Req
'            Integra800_QCRes_Req
'
'        Else
'            Integra800_OrderID_Req
''            Integra800_Res_Req
''            Integra800_QCRes_Req
'        End If
''            Integra800_QCRes_Req
'    ElseIf Host_BC = "10" Then
'
'        If vasCode.DataRowCnt < 1 Then
'            Integra800_OrderID_Req
'        End If
''    Else
''        Integra800_OrderID_Req
''        Integra800_QCRes_Req
'    End If
'End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    
    Dim lsTempBarCode As String
    Dim lsPID As String
    Dim lsPname As String
    Dim lsSex As String
    Dim lsAge As String
    
    On Error GoTo erroCheck
    
    '샘플번호에 해당 하는 검사결과 Local Databse에서 가져오기
        
    ClearSpread vasRes
    vasRes.MaxRows = 0
    
    lsID = Trim(GetText(vasID, Row, colBarCode))
    
    
    SQL = "select equipcode, examcode, examname, resvalue, result, seqno, examdate, examtime , refflag, '', panicflag, RESFLAG, deltaflag" & vbCrLf & _
          "FROM pat_res " & vbCrLf & _
          "WHERE Barcode = '" & Trim(GetText(vasID, Row, colBarCode)) & "' " & vbCrLf & _
          "GROUP BY equipcode, examcode, examname, resvalue, result, seqno, examdate, examtime , refflag,  panicflag, RESFLAG, deltaflag " & vbCrLf & _
          "  order by seqno, equipcode"
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
   
'         "  AND EXAMDATE = '" & Trim(Format(dtpToday, "yyyymmdd")) & "'" & vbCrLf & _

    
    For i = 1 To vasRes.DataRowCnt
        If Trim(GetText(vasRes, i, colEquipExam)) = "YEA" Then
            If Trim(GetText(vasRes, i, colResValue)) = "Found" Then
                SetBackColor vasRes, i, i, colEquipExam, colORDERFLAG, 200, 100, 100
            End If
        End If
    Next i
    
    If Trim(GetText(vasID, Row, colBarCode)) = "" Then Exit Sub
    
    'cmdImage.Visible = False
    
    Exit Sub
    
erroCheck:
    'Save_Raw_Data gImagePath & "\*" & GetText(vasID, Row, colBarCode) & "*" & vbCrLf & Err.Description
    
End Sub

'''Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
'''
'''On Error GoTo erroCheck
'''
'''    If Trim(GetText(vasID, Row, colBarCode)) = "" Then Exit Sub
'''
'''    lblRowNum.Caption = Row
'''    If Dir(gImagePath & "\cobas_6500_ResultReport_" & GetText(vasID, Row, colBarCode) & "_*", vbDirectory) <> "" Then
'''        cmdImage.Visible = True
'''    Else
'''        cmdImage.Visible = False
'''    End If
'''
'''    Exit Sub
'''
'''erroCheck:
'''    Save_Raw_Data gImagePath & "\*" & GetText(vasID, Row, colBarCode) & "*" & vbCrLf & Err.Description
'''
'''End Sub

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sSendFlag As String
    Dim sResFlag As String
    
    sResFlag = ""
    sExamDate = ""
    sExamDate = Trim(GetText(vasRes, asRow2, colResDate))
    sExamTime = Trim(GetText(vasRes, asRow2, colResTime))
    sResFlag = Trim(GetText(vasRes, asRow2, colResTime))
    If Trim(sExamDate) = "" Then
        sExamDate = Format(Date, "yyyymmdd")
    End If
    
    If Trim(GetText(vasID, asRow1, colState)) = "미접수" Then
        asSend = "5"
    ElseIf Trim(GetText(vasID, asRow1, colState)) = "오더없음" Then
        asSend = "6"
    End If
    
    If strORDERType = "" Then
        '오더타입 기억(DISKNO)
        SQL = ""
        SQL = SQL & vbCrLf & "SELECT DISKNO"
        SQL = SQL & vbCrLf & "  FROM PAT_RES "
        SQL = SQL & vbCrLf & " WHERE BARCODE = '" & gSpecID & "'"
        SQL = SQL & vbCrLf & "   AND DISKNO <> ''"
        res = db_select_Col(gLocal, SQL)
        
        strORDERType = Trim(gReadBuf(0))
    End If
    
    SQL = "select examcode FROM pat_res " & vbCrLf & _
          "WHERE " & vbCrLf & _
          "  equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' "
    'SQL = SQL & vbCrLf & " AND POSNO = '" & Trim(GetText(vasID, asRow1, colRackPos)) & "'"
    res = db_select_Row(gLocal, SQL)
    
    If res > 0 Then
        SQL = "update pat_res set resvalue = '" & Trim(GetText(vasRes, asRow2, colResValue)) & "', " & vbCrLf & _
              "result = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
              "sendflag = '" & asSend & "', " & vbCrLf & _
              "refflag = '" & Trim(GetText(vasRes, asRow2, colAFLAG)) & "', " & vbCrLf & _
              "panicflag = '" & Trim(GetText(vasRes, asRow2, colPFLAG)) & "', " & vbCrLf & _
              "examdate = '" & sExamDate & "', examtime = '" & sExamTime & "', " & vbCrLf & _
              "posno = '" & Trim(GetText(vasID, asRow1, colRackPos)) & "' , psex = '" & Trim(GetText(vasID, asRow1, colPSex)) & "' " & vbCrLf & _
              ",EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "',examname = '" & Trim(GetText(vasRes, asRow2, colExamName)) & "'" & vbCrLf & _
              ",resflag = ' " & Trim(GetText(vasRes, asRow2, colRESFLAG)) & "'" & vbCrLf & _
              ",receno = ' " & Trim(GetText(vasID, asRow1, colWkNo)) & "'" & vbCrLf & _
              ",deltaflag = ' " & Trim(GetText(vasRes, asRow2, colORDERFLAG)) & "'" & vbCrLf & _
              ",EXAMUID = ' " & Trim(GetText(vasID, asRow1, colDept)) & "'" & vbCrLf & _
              ",DISKNO = ' " & strORDERType & "'" & vbCrLf & _
              "WHERE  " & vbCrLf & _
              "  equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "'" & vbCrLf & _
              "  AND barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' "
        'SQL = SQL & vbCrLf & " AND POSNO = '" & Trim(GetText(vasID, asRow1, colRackPos)) & "'"
        res = SendQuery(gLocal, SQL)
        
    Else
        SQL = "insert into pat_res(examdate, equipno, barcode, equipcode, examcode, " & vbCrLf & _
              "refflag, sendflag, seqno, examname, resvalue, " & vbCrLf & _
              "result, examtime, pid, pname, panicflag, posno, psex, deltaflag, receno, EXAMUID,DISKNO) " & vbCrLf & _
              "values('" & sExamDate & "', '" & gEquip & "', '" & Trim(GetText(vasID, asRow1, colBarCode)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colEquipExam)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colAFLAG)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', '" & Trim(GetText(vasRes, asRow2, colResValue)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
              "'" & sExamTime & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & _
              Trim(GetText(vasRes, asRow2, colPFLAG)) & "', '" & Trim(GetText(vasID, asRow1, colRackPos)) & "' , '" & _
              Trim(GetText(vasID, asRow1, colPSex)) & "', '" & _
              Trim(GetText(vasRes, asRow2, colORDERFLAG)) & "', '" & _
              Trim(GetText(vasID, asRow1, colWkNo)) & "','" & Trim(GetText(vasID, asRow1, colDept)) & "','" & strORDERType & "') "
        res = SendQuery(gLocal, SQL)
    End If
    
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

Private Sub vasID_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    
    vasID_Click NewCol, NewRow
    
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If

    PopupMenu mnuPop
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
    
    ClearSpread vasLISTRES
    vasLISTRES.MaxRows = 0
    
    lsID = Trim(GetText(vasList, Row, colBarCode))
        SQL = "select A.equipcode, A.examcode, A.examname, A.resvalue, A.result, B.RESseqno, A.examdate, A.examtime , A.refflag, '', A.panicflag, A.RESFLAG, A.deltaflag" & vbCrLf & _
          "FROM pat_res A , EQUIPEXAM B " & vbCrLf & _
          "WHERE A.EQUIPCODE = B.EQUIPCODE " & vbCrLf & _
          "  AND A.Barcode = '" & Trim(GetText(vasList, Row, colBarCode)) & "'" & vbCrLf & _
          "GROUP BY A.equipcode, A.examcode, A.examname, A.resvalue, A.result, B.RESseqno, A.examdate, A.examtime , A.refflag, '', A.panicflag, A.RESFLAG, A.deltaflag" & vbCrLf & _
          "  order by B.RESseqno, A.equipcode"

    res = db_select_Vas(gLocal, SQL, vasLISTRES)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    lblInfo1.Caption = lsID & " / " & Trim(GetText(vasList, Row, colPName))
    lblInfo2.Caption = Trim(GetText(vasList, Row, colWkNo)) '& " / " & Trim(GetText(vasList, Row, colBarCode))
    
    '"  AND EXAMDATE = '" & Trim(Format(dtpExamDate, "yyyymmdd")) & "'" & vbCrLf & _

    If vasLISTRES.DataRowCnt > 12 Then
        InsertRow vasLISTRES, 13
        SetBackColor vasLISTRES, 13, 13, 1, vasLISTRES.MaxCols, 100, 150, 255
        SetText vasLISTRES, "Microscopy", 13, colExamName
    End If
    
    
    
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col <> 12 Then Exit Sub
    If Trim(GetText(vasList, Row, Col)) = "" Then Exit Sub
    
    
    
End Sub

Private Sub vasList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If NewRow = -1 Or NewCol = -1 Then Exit Sub
    
    vasList_Click NewCol, NewRow
End Sub

Private Sub vasLISTRES_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    
End Sub

Private Sub vasres_rightclick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim vasIDRow As Integer
    Dim VasResRow As Integer
    
    vasIDRow = vasID.ActiveRow
    VasResRow = vasRes.ActiveRow
    If vasIDRow < 1 Or vasIDRow > vasID.DataRowCnt Then
        Exit Sub
    End If
    If VasResRow < 1 Or VasResRow > vasRes.DataRowCnt Then
        Exit Sub
    End If
    
    PopupMenu mnuPop

End Sub

Private Sub subDel_Click()
    Dim i As Long
    Dim vasIDRow As Integer
    Dim VasResRow As Integer
    Dim X As Long
    Dim j As Long
    Dim c, r, c2, r2

    vasIDRow = vasID.ActiveRow
    VasResRow = vasRes.ActiveRow
    If vasIDRow < 1 Or vasIDRow > vasID.DataRowCnt Then
        Exit Sub
    End If
    If VasResRow < 1 Or VasResRow > vasRes.DataRowCnt Then
        Exit Sub
    End If

    If vasRes.IsBlockSelected Or vasRes.SelectionCount Then

        vasRes.BlockMode = True
'        db_BeginTran gLocal
        
        For X = 0 To vasRes.SelectionCount - 1
            vasRes.GetSelection X, c, r, c2, r2
            vasRes.Col = c
            vasRes.Col2 = c2
            vasRes.Row = r
            vasRes.Row2 = r2
            If IsNumeric(r) = True And IsNumeric(r2) = True Then
                If CInt(r) > 0 And CInt(r2) > 0 Then
                    For j = r To r2
                        SQL = "Delete from pat_res where barcode = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
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
    
    vasID_Click colBarCode, vasIDRow
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

'WinSock Control ==============================================================================================================
Public Sub WinSock_Listen(argWinSock As Winsock)
    Dim sWinSockPort As String
    
    On Error GoTo ErrorCheck

    sWinSockPort = gSetup.gPort
    
    
    If sWinSockPort = "0" Or IsNumeric(sWinSockPort) = False Then
        Exit Sub
    End If
    
    If argWinSock.State <> sckClosed Then
        argWinSock.Close
    End If
    
    'argWinSock.RemoteHost = "162.132.241.155"
    argWinSock.LocalPort = sWinSockPort
    argWinSock.Listen
    'argWinSock.Connect "162.132.241.115", "6500"
    '    argWinSock.RemoteHost = "162.132.241.115"
    'argWinSock.RemotePort = "6500"
    'argWinSock.Connect
    
'''    If EquipNum = 1 Then
'''        lblConnect1.Caption = "연결 대기중..."
'''    Else
'''        lblConnect2.Caption = "연결 대기중..."
'''    End If
    
    Exit Sub
    
ErrorCheck:
    
    MsgBox gEquipName & "이(가) 정상적으로 실행되지 않았습니다. " & vbCrLf & "다시 한번 실행해 주세요." & vbCrLf & "ERROR : " & Err.Number & " - " & Err.Description
    KillProcess "IF_" & gEquipName & ".exe"
    
    
    'ExecuteProcess "IF_Cobas6500.exe"
    
    
End Sub

Private Sub Winsock1_Close()
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    Winsock1.LocalPort = gSetup.gPort
    Winsock1.Listen
    
    'Interface Program
    frmInterface.Caption = "Interface Program - " & "연결 대기중..."
'''    lblConnect1.Caption = "연결 대기중..."
    
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    
    Winsock1.Accept requestID
    
    frmInterface.Caption = "Interface Program - " & "연결[" & requestID & "]" & Winsock1.RemoteHostIP
'''    lblConnect1.Caption = "연결[" & requestID & "]" & Winsock1.RemoteHostIP
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim sTmp As String
    Dim strSendData
    Dim strResFlag
    Dim sSigFlag As String
    Dim sStemp As String
    
    Dim arrSIGNAL
    Dim intX    As Integer
    
On Error GoTo errCHECK

    Winsock1.GetData sTmp
    
    'Save_Raw_Data2 "[RX" & Format(time, "hh:mm:ss") & "]" & sTmp
    
    If InStr(1, sTmp, chrACK) = 0 Then
        txtWinSockBuff = txtWinSockBuff & sTmp
    End If
    
    If InStr(1, sTmp, chrENQ) > 0 Then
        blTimerChk = False
        
        Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & txtWinSockBuff
        txtWinSockBuff = ""
        
        Call Winsock1.SendData(chrACK)
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
        Exit Sub
    End If
    
    If InStr(1, sTmp, chrLF) > 0 Then
        
        Call Winsock1.SendData(chrACK)
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & chrACK
        
        If InStr(1, sTmp, "2Q") > 0 Then
            strQMode = "Q"
        End If
        Exit Sub
    End If
    
    If InStr(1, sTmp, chrACK) > 0 Then
        
        If gOrderMsg(0) = "" And gOrderMsg(2) = "" And gOrderMsg(3) = "" And gOrderMsg(4) <> "" Then
                Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(4)
                blTimerChk = True
                blTimerTerminate = True
                Call Winsock1.SendData(gOrderMsg(4))
                gOrderMsg(4) = ""
                
                'If GetText(vasOrderTest, 2, 1) = "" Then
                    DeleteRow vasOrderTest, 1, 1
                'End If
                
        ElseIf gOrderMsg(0) = "" And gOrderMsg(2) = "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(3)
            Call Winsock1.SendData(gOrderMsg(3))
            gOrderMsg(3) = ""
            
        ElseIf gOrderMsg(0) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(2)
            Call Winsock1.SendData(gOrderMsg(2))
            gOrderMsg(2) = ""
            
'        ElseIf gOrderMsg(0) = "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
'            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(1)
'            Call Winsock1.SendData(gOrderMsg(1))
'            gOrderMsg(1) = ""
            
        ElseIf gOrderMsg(0) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(0)
            Call Winsock1.SendData(gOrderMsg(0))
            gOrderMsg(0) = ""
            
        '끝나는 신호가 없을경우
        ElseIf gOrderMsg(0) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(0)
            Call Winsock1.SendData(gOrderMsg(0))
            gOrderMsg(0) = ""
'        ElseIf gOrderMsg(0) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
'            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(1)
'            Call Winsock1.SendData(gOrderMsg(1))
'            gOrderMsg(1) = ""
        ElseIf gOrderMsg(0) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
            Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & gOrderMsg(2)
            Call Winsock1.SendData(gOrderMsg(2))
            gOrderMsg(2) = ""
            DeleteRow vasOrderTest, 1, 1
            FN_OrderSend
        End If
        
        
''        If gOrderMsg(0) = "" And gOrderMsg(1) = "" And gOrderMsg(2) = "" And gOrderMsg(3) = "" And gOrderMsg(4) <> "" Then
''                Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(4)
''                blTimerChk = True
''                blTimerTerminate = True
''                Call Winsock1.SendData(gOrderMsg(4))
''                gOrderMsg(4) = ""
''
''                'If GetText(vasOrderTest, 2, 1) = "" Then
''                    DeleteRow vasOrderTest, 1, 1
''                'End If
''
''        ElseIf gOrderMsg(0) = "" And gOrderMsg(1) = "" And gOrderMsg(2) = "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(3)
''            Call Winsock1.SendData(gOrderMsg(3))
''            gOrderMsg(3) = ""
''
''        ElseIf gOrderMsg(0) = "" And gOrderMsg(1) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(2)
''            Call Winsock1.SendData(gOrderMsg(2))
''            gOrderMsg(2) = ""
''
''        ElseIf gOrderMsg(0) = "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(1)
''            Call Winsock1.SendData(gOrderMsg(1))
''            gOrderMsg(1) = ""
''
''        ElseIf gOrderMsg(0) <> "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) <> "" And gOrderMsg(4) <> "" Then
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(0)
''            Call Winsock1.SendData(gOrderMsg(0))
''            gOrderMsg(0) = ""
''
''        '끝나는 신호가 없을경우
''        ElseIf gOrderMsg(0) <> "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(0)
''            Call Winsock1.SendData(gOrderMsg(0))
''            gOrderMsg(0) = ""
''        ElseIf gOrderMsg(0) = "" And gOrderMsg(1) <> "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(1)
''            Call Winsock1.SendData(gOrderMsg(1))
''            gOrderMsg(1) = ""
''        ElseIf gOrderMsg(0) = "" And gOrderMsg(1) = "" And gOrderMsg(2) <> "" And gOrderMsg(3) = "" And gOrderMsg(4) = "" Then
''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & gOrderMsg(2)
''            Call Winsock1.SendData(gOrderMsg(2))
''            gOrderMsg(2) = ""
''            DeleteRow vasOrderTest, 1, 1
''            FN_OrderSend
''        End If
        
        
'
'
'        Call Winsock1.SendData(GetText(vasOrder, 1, 1))
'
'        Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & GetText(vasOrder, 1, 1)
'
'        If InStr(1, GetText(vasOrder, 1, 1), chrEOT) > 0 Then
'            blTimerChk = True
'        End If
'
'        DeleteRow vasOrder, 1, 1
        
        
        Exit Sub
    End If
    
    If InStr(1, sTmp, chrEOT) > 0 Then
        Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & "]" & txtWinSockBuff
        
        sSigFlag = Cobas8000(txtWinSockBuff)
        txtWinSockBuff = ""

        If strQMode = "Q" Then
            'Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrENQ
            'Call Winsock1.SendData(chrENQ)
            
            strQMode = ""
        End If
        
        If vasOrderTest.DataRowCnt > 0 Then
            intTimer = 0
            
        End If
        blTimerChk = True
'''            If vasOrder.DataRowCnt > 0 Then
'''                intTimer = 0
'''            End If
        
'''        If vasOrder.DataRowCnt > 0 Then
'''            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrENQ
'''            Call Winsock1.SendData(chrENQ)
'''        End If
        
''
'        If vasOrder.DataRowCnt > 0 Then
'            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrENQ
'            Call Winsock1.SendData(chrENQ)
'        End If
        blTimerChk = True
        
        Exit Sub
    End If
    
    Exit Sub
'    If InStr(1, sTmp, chrENQ) > 0 Then
'
'        txtBuff.Text = txtBuff.Text & sTmp
'        gOrderMessage = ""
'        gOrderCnt = 1
'        comSend = "stENQ"
'
'        arrSIGNAL = Split(txtBuff.Text, chrVT)
'
'        For intX = 0 To UBound(arrSIGNAL)
'            'sSigFlag = Cobas8000_HL7(CStr(arrSIGNAL(intX)))
'            sSigFlag = Cobas8000(CStr(arrSIGNAL(intX)))
'        Next intX
'
'
'        txtBuff.Text = ""
'        If sSigFlag = "Q" Then
'            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrENQ
'            Winsock1.SendData chrENQ
'        End If
'    Else
'        txtBuff.Text = txtBuff.Text & sTmp
'    End If
    Exit Sub
errCHECK:
    Save_Raw_Data "[ERR WinsoCK]" & Format(Time, "hh:mm:ss")
   
'    If InStr(1, sTmp, chrENQ) > 0 Then
'        txtBuff.text = sTmp
'
'        Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrACK
'        Winsock1.SendData chrACK
'    End If
'
'    If InStr(1, sTmp, chrLF) > 0 Then
'        txtBuff.text = txtBuff.text & sTmp
'        Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrACK
'        Winsock1.SendData chrACK
'    End If
'
'    If InStr(1, sTmp, chrEOT) > 0 Then
'
'        txtBuff.text = txtBuff.text & sTmp
'        gOrderMessage = ""
'        gOrderCnt = 1
'        comSend = "stENQ"
'
'        sSigFlag = Cobas8000(txtBuff.text)
'        If sSigFlag = "Q" Then
'
'            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrENQ
'            Winsock1.SendData chrENQ
'        End If
'
'    End If
'
'    If InStr(1, sTmp, chrACK) > 0 Then
'        Save_Raw_Data "[RX" & Format(time, "hh:mm:ss") & "]" & chrACK
'        If comSend = "stENQ" Then
'            sStemp = SendOrder
'            Winsock1.SendData sStemp
'            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & sStemp
'
'        ElseIf comSend = "stOrder" Then
'            Winsock1.SendData chrEOT
'            Save_Raw_Data "[TX" & Format(time, "hh:mm:ss") & "]" & chrEOT
'
'        End If
'
'    End If
    
    
End Sub

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
    Dim j As Integer
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
    lsData = Replace(lsData, chrETX, "")
    
    If lsData = vbCrLf Then
        Exit Function
    End If
    
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
'''    i = InStr(1, lsData, chrETB)
'''
'''    While i > 0
'''        lsData = Mid(lsData, 1, i - 4) & Mid(lsData, i + 1)
'''        i = InStr(1, lsData, chrETB)
'''    Wend
    
    
    'i = InStr(1, lsData, Chr(13))
    
    Dim arrData
    
    arrData = Split(lsData, Chr(13))
    
    For j = 0 To UBound(arrData) - 1
        lsTemp = arrData(j)
        
        If lsTemp <> "" Then
            CobasProg lsTemp
            lsTemp = ""
        End If
        
    Next j
    
'    Do While i > 0
'        lsTemp = Mid(lsData, 1, i - 1)
'        lsData = Mid(lsData, i + 1)
'
'
'
'        Select Case Left(lsTemp, 1)
'        Case "Q"
'            lsMSGflag = "Q"
'        Case "O"
'            lsMSGflag = "O"
'        End Select
'
'        CobasProg lsTemp
'
'        i = InStr(1, lsData, chrCR)
'    Loop
    
    Cobas8000 = lsMSGflag
End Function

Function ADP_CHECK(argBarcode As String, argSEX As String, argEXAMCODE As String, argResult As String) As String
    
'         select top 1 pJUDGCHRL  --crr하한
'                    , pJUDGCHRH --crr상한
'                    , ccrmin --ccr하한
'                    , ccrmax --ccr상한
'                    , PANICMIN --패닉하한
'                    , PANICMAX --패닉상한
'                    , rvalmin --판정하한
'                    , RVALMAX --판정상한
'                    , judgchrh --판정상한 판정문자(판정높을경우 만약문자가 없을경우 'H')
'                    , judgchrh --판정하한 판정문자(판정높을경우 만약문자가 없을경우 'L')
'                    , CharRval --판정문자
'                    , b.deltaflag
'                    , b.deltaval from shinbase..intrsltcd a , shinbase..testcd b
'         Where a.testcd_cd = b.cd
'           and a.testcd_cd = '00178' --검사코드
'           and ( a.sex = '3' or a.sex = '2') --성별 '3' 공통 '1'남 '2' 여
'           and convert(float,a.startage) <= '10'
'           and convert(float,a.endage) >= '10'
'           and a.appdd <= '20121228'
'           and b.judgyn <> 'N'
'           and a.del_yn <> 'Y'
'         order by a.appdd desc
    
    ADP_CHECK = "/"
    
    Dim strAGE  As String
    
    Dim pJUDGCHRL As String
    Dim pJUDGCHRH As String
    Dim ccrmin As String
    Dim ccrmax As String
    Dim PANICMIN As String
    Dim PANICMAX As String
    Dim rvalmin As String
    Dim RVALMAX As String
    Dim judgchrh As String
    Dim CharRval As String
    Dim MAXREF As String
    Dim MINREF As String
    'MAXREF, MINREF
    
    Dim AFLAG As String
    Dim PFLAG As String
    
On Error GoTo ErrHandle
    argSEX = ""
    AFLAG = ""
    PFLAG = ""
    If IsNumeric(argResult) = False Then Exit Function
    
    If Len(argBarcode) <> 15 Then Exit Function
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT ACPTDD, ACPTNO, NM, RGSTNO, SEX, AGE, "
    SQL = SQL & vbCrLf & "      (select NM from shinbase..custcd c where a.CUSTCD_CD = c.CD)"
    SQL = SQL & vbCrLf & "  from shinrslt..acptinfo a"
    SQL = SQL & vbCrLf & " where acptdd = '" & "20" & Mid(argBarcode, 1, 6) & "' "
    SQL = SQL & vbCrLf & "   and acptno = '" & Mid(argBarcode, 7, 7) & "'"

    res = db_select_Col(gServer, SQL)
    
    If res > 0 Then
        strAGE = Val(gReadBuf(5) & "")
        argSEX = Trim(gReadBuf(4) & "")
    End If
    
    If strAGE = "" Then strAGE = "0"
    If argSEX = "" Then argSEX = "3"
    
    
    
    Dim i As Integer
    For i = 0 To 20
        gReadBuf(i) = ""
    Next i
    
    '0.5 25.0    1.5 12.0    1.5 12.0    3.5 7.2
    SQL = ""
    SQL = SQL & vbCrLf & "select top 1 pJUDGCHRL "
    SQL = SQL & vbCrLf & "           , pJUDGCHRH "
    SQL = SQL & vbCrLf & "           , ccrmin "
    SQL = SQL & vbCrLf & "           , ccrmax "
    SQL = SQL & vbCrLf & "           , PANICMIN "
    SQL = SQL & vbCrLf & "           , PANICMAX "
    SQL = SQL & vbCrLf & "           , rvalmin "
    SQL = SQL & vbCrLf & "           , RVALMAX "
    SQL = SQL & vbCrLf & "           , judgchrh "
    SQL = SQL & vbCrLf & "           , CharRval "
    SQL = SQL & vbCrLf & "           , b.deltaflag "
    SQL = SQL & vbCrLf & "           , b.deltaval "
    SQL = SQL & vbCrLf & "           , MAXREF "
    SQL = SQL & vbCrLf & "           , MINREF "
    SQL = SQL & vbCrLf & "  from shinbase..intrsltcd a , shinbase..testcd b"
    SQL = SQL & vbCrLf & " Where a.testcd_cd = b.cd "
    SQL = SQL & vbCrLf & "   and a.testcd_cd = '" & argEXAMCODE & "'"
    SQL = SQL & vbCrLf & "   and ( a.sex = '3' or a.sex = '" & argSEX & "' )"
    SQL = SQL & vbCrLf & "   and convert(float,a.startage) <= '" & strAGE & "'"
    SQL = SQL & vbCrLf & "   and convert(float,a.endage) >= '" & strAGE & "'"
    SQL = SQL & vbCrLf & "   and a.appdd <= '" & "20" & Mid(argBarcode, 1, 6) & "' "
    SQL = SQL & vbCrLf & "   and b.judgyn <> 'N'"
    SQL = SQL & vbCrLf & "   and a.del_yn <> 'Y'"
    SQL = SQL & vbCrLf & " order by a.appdd "
    
    
    res = db_select_Col(gServer, SQL)
    
    pJUDGCHRL = gReadBuf(0)
    pJUDGCHRH = gReadBuf(1)
    ccrmin = gReadBuf(2)
    ccrmax = gReadBuf(3)
    
    PANICMIN = gReadBuf(4)
    PANICMAX = gReadBuf(5)
    
    rvalmin = gReadBuf(6)
    RVALMAX = gReadBuf(7)
    judgchrh = gReadBuf(8)
    
    CharRval = gReadBuf(9)
    
    MAXREF = gReadBuf(12)
    MINREF = gReadBuf(13)
    
    'Save_Raw_Data "[ADP CHECK]" & gReadBuf(0) & "/" & gReadBuf(1) & "/" & gReadBuf(2) & "/" & gReadBuf(3) & "/" & gReadBuf(4) & "/" & gReadBuf(5) & "/" & gReadBuf(6) & "/" & _
                                  gReadBuf(7) & "/" & gReadBuf(8) & "/" & gReadBuf(9) & "/" & gReadBuf(9) & "/" & gReadBuf(9) & "/" & gReadBuf(10) & "/" & gReadBuf(11) & "/" & _
                                  gReadBuf(12) & "/" & gReadBuf(13)
    
    'Save_Raw_Data "[ADP CHECK]" & SQL
    
    '/A FLAG
    If rvalmin = "0" And RVALMAX = "0" Then
        If IsNumeric(MINREF) = True And IsNumeric(MAXREF) = True Then
            If Val(argResult) < Val(MINREF) Then
                If Trim(judgchrh) <> "" Then
                    AFLAG = judgchrh
                Else
                    AFLAG = "L"
                End If
            ElseIf Val(argResult) > Val(MAXREF) Then
                If Trim(judgchrh) <> "" Then
                    AFLAG = judgchrh
                Else
                    AFLAG = "H"
                End If
            End If
            
        ElseIf IsNumeric(MINREF) = True And IsNumeric(MAXREF) = False Then
            If Val(argResult) < Val(MINREF) Then
                If Trim(judgchrh) <> "" Then
                    AFLAG = judgchrh
                Else
                    AFLAG = "L"
                End If
            End If
        ElseIf IsNumeric(MINREF) = False And IsNumeric(MAXREF) = True Then
            If Val(argResult) > Val(MAXREF) Then
                If Trim(judgchrh) <> "" Then
                    AFLAG = judgchrh
                Else
                    AFLAG = "H"
                End If
            End If
        End If
    Else
        If IsNumeric(rvalmin) = True And IsNumeric(RVALMAX) = True Then
            If Val(argResult) < Val(rvalmin) Then
                If Trim(judgchrh) <> "" Then
                    AFLAG = judgchrh
                Else
                    AFLAG = "L"
                End If
            ElseIf Val(argResult) > Val(RVALMAX) Then
                If Trim(judgchrh) <> "" Then
                    AFLAG = judgchrh
                Else
                    AFLAG = "H"
                End If
            End If
            
        ElseIf IsNumeric(rvalmin) = True And IsNumeric(RVALMAX) = False Then
            If Val(argResult) < Val(rvalmin) Then
                If Trim(judgchrh) <> "" Then
                    AFLAG = judgchrh
                Else
                    AFLAG = "L"
                End If
            End If
        ElseIf IsNumeric(rvalmin) = False And IsNumeric(RVALMAX) = True Then
            If Val(argResult) > Val(RVALMAX) Then
                If Trim(judgchrh) <> "" Then
                    AFLAG = judgchrh
                Else
                    AFLAG = "H"
                End If
            End If
        End If
    End If
    
    
    '/P FLAG
    If PANICMIN = "0" And PANICMAX = "0" Then
    Else
        If IsNumeric(PANICMIN) = True And IsNumeric(PANICMAX) = True Then
            If Val(argResult) < Val(PANICMIN) Then
                PFLAG = "P"
            ElseIf Val(argResult) > Val(PANICMAX) Then
                PFLAG = "P"
            End If
        ElseIf IsNumeric(PANICMIN) = True And IsNumeric(PANICMAX) = False Then
            If Val(argResult) < Val(PANICMIN) Then
                PFLAG = "P"
            End If
        ElseIf IsNumeric(PANICMIN) = False And IsNumeric(PANICMAX) = True Then
            If Val(argResult) > Val(PANICMAX) Then
                PFLAG = "P"
            End If
        End If
    End If
    ADP_CHECK = AFLAG & "/" & PFLAG
    
    Exit Function
    
ErrHandle:
    ADP_CHECK = "/"
End Function

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'''    lblConnect1.Caption = "[Error]" & Number & " : " & Description
End Sub

Function SetVasColor(argVasSp As vaSpread, argRow As Integer, argBarcode As String)
    Dim strVasName As String
    Dim strFlag As String
    Dim intAcnt As Integer
    strVasName = argVasSp.Name
    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT RESFLAG"
    SQL = SQL & vbCrLf & "  FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & argBarcode & "'"
    SQL = SQL & vbCrLf & " GROUP BY RESFLAG"
    res = db_select_Col(gLocal, SQL)
    strFlag = Trim(gReadBuf(0))
    intAcnt = 0
    If strFlag = "" Then Exit Function
    
    '중요도 순서(역순)
    If InStr(1, strFlag, "C") > 0 Then
        intAcnt = intAcnt - 1
        SetBackColor argVasSp, argRow, argRow, colRackPos, colPSex, 255, 187, 0
    End If
    
    If InStr(1, strFlag, "D") > 0 Then
        intAcnt = intAcnt - 1
        SetBackColor argVasSp, argRow, argRow, colRackPos, colPSex, 128, 65, 217
    End If
    
    If InStr(1, strFlag, "B") > 0 Then
        intAcnt = intAcnt - 1
        SetBackColor argVasSp, argRow, argRow, colRackPos, colPSex, 255, 255, 90
    End If
    
    If InStr(1, strFlag, "A") > 0 Then
        intAcnt = intAcnt + 1
        
        If intAcnt > 0 Then
            SetBackColor argVasSp, argRow, argRow, colRackPos, colDept, 54, 138, 255
        Else
            SetBackColor argVasSp, argRow, argRow, colDept, colDept, 54, 138, 255
        End If
    End If
    
    
End Function


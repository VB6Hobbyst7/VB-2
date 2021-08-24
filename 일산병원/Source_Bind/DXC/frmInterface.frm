VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   0  '없음
   Caption         =   " DxC Interface "
   ClientHeight    =   10470
   ClientLeft      =   2790
   ClientTop       =   1125
   ClientWidth     =   15045
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmInterface.frx":030A
   ScaleHeight     =   10470
   ScaleMode       =   0  '사용자
   ScaleWidth      =   15045
   Begin VB.Frame fraResFlag 
      Height          =   3480
      Left            =   15840
      TabIndex        =   46
      Top             =   5760
      Visible         =   0   'False
      Width           =   8385
      Begin FPSpread.vaSpread vasResMemo 
         Height          =   2850
         Left            =   45
         TabIndex        =   47
         Top             =   540
         Width           =   8250
         _Version        =   393216
         _ExtentX        =   14552
         _ExtentY        =   5027
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
         MaxCols         =   1
         SpreadDesigner  =   "frmInterface.frx":058D
      End
      Begin VB.Label lblFlagPname 
         Caption         =   "1234567890ab"
         Height          =   225
         Left            =   4215
         TabIndex        =   51
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label Label5 
         Caption         =   "환자명 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3165
         TabIndex        =   50
         Top             =   225
         Width           =   945
      End
      Begin VB.Label lblFlagBarcode 
         Caption         =   "1234567890ab"
         Height          =   165
         Left            =   1620
         TabIndex        =   49
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label Label7 
         Caption         =   "바코드번호 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   48
         Top             =   225
         Width           =   1380
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   2655
      Left            =   360
      TabIndex        =   27
      Top             =   6900
      Visible         =   0   'False
      Width           =   8205
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   870
         Left            =   6795
         TabIndex        =   54
         Top             =   315
         Width           =   420
         _Version        =   393216
         _ExtentX        =   741
         _ExtentY        =   1535
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
         SpreadDesigner  =   "frmInterface.frx":1D73
      End
      Begin FPSpread.vaSpread vasPatList 
         Height          =   330
         Left            =   5040
         TabIndex        =   53
         Top             =   2070
         Width           =   915
         _Version        =   393216
         _ExtentX        =   1614
         _ExtentY        =   582
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
         SpreadDesigner  =   "frmInterface.frx":7B83
      End
      Begin FPSpread.vaSpread vasOrderBuf 
         Height          =   600
         Left            =   855
         TabIndex        =   52
         Top             =   1170
         Width           =   600
         _Version        =   393216
         _ExtentX        =   1058
         _ExtentY        =   1058
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
         SpreadDesigner  =   "frmInterface.frx":7D9B
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   7290
         TabIndex        =   44
         Top             =   270
         Width           =   825
         _Version        =   393216
         _ExtentX        =   1455
         _ExtentY        =   1614
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
         SpreadDesigner  =   "frmInterface.frx":7FB3
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   42
         Top             =   1980
         Width           =   1215
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
         Height          =   435
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   33
         Top             =   1380
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1125
      End
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
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1620
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   30
         Top             =   2025
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   3600
         Style           =   1  '그래픽
         TabIndex        =   29
         Top             =   1980
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   960
         Left            =   6705
         TabIndex        =   28
         Top             =   1350
         Visible         =   0   'False
         Width           =   1335
         Begin MSCommLib.MSComm MSComm1 
            Left            =   135
            Top             =   300
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
            InputLen        =   1
            RThreshold      =   1
            RTSEnable       =   -1  'True
            EOFEnable       =   -1  'True
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   720
            Top             =   270
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   4980
         TabIndex        =   34
         Top             =   180
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
         SpreadDesigner  =   "frmInterface.frx":81CB
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   900
         Left            =   1485
         TabIndex        =   61
         Top             =   180
         Width           =   1800
         _Version        =   393216
         _ExtentX        =   3175
         _ExtentY        =   1587
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
         SpreadDesigner  =   "frmInterface.frx":83E3
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   735
         Left            =   1800
         TabIndex        =   62
         Top             =   1125
         Width           =   1680
         _Version        =   393216
         _ExtentX        =   2963
         _ExtentY        =   1296
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
         SpreadDesigner  =   "frmInterface.frx":85FB
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   810
         Left            =   3285
         TabIndex        =   63
         Top             =   225
         Width           =   1275
         _Version        =   393216
         _ExtentX        =   2249
         _ExtentY        =   1429
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
         SpreadDesigner  =   "frmInterface.frx":8813
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   4860
         TabIndex        =   36
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   5700
         TabIndex        =   35
         Top             =   1410
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   2085
      Left            =   15360
      TabIndex        =   24
      Top             =   2880
      Visible         =   0   'False
      Width           =   9465
      Begin FPSpread.vaSpread vasTest 
         Height          =   675
         Left            =   60
         TabIndex        =   66
         Top             =   720
         Width           =   1695
         _Version        =   393216
         _ExtentX        =   2990
         _ExtentY        =   1191
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
         SpreadDesigner  =   "frmInterface.frx":8A2B
      End
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   1260
         TabIndex        =   25
         Top             =   240
         Width           =   8160
         _Version        =   393216
         _ExtentX        =   14393
         _ExtentY        =   2725
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
         MaxCols         =   9
         SpreadDesigner  =   "frmInterface.frx":8C43
      End
      Begin FPSpread.vaSpread vasPrintBuf 
         Height          =   1245
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1065
         _Version        =   393216
         _ExtentX        =   1879
         _ExtentY        =   2196
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
         SpreadDesigner  =   "frmInterface.frx":A6BC
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   9240
      Left            =   75
      TabIndex        =   6
      Top             =   840
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   16298
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "진행상태"
      TabPicture(0)   =   "frmInterface.frx":A8D4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":A8F0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   8775
         Left            =   -74820
         TabIndex        =   17
         Top             =   360
         Width           =   14625
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   8460
            TabIndex        =   37
            Top             =   630
            Width           =   6015
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   43
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Left            =   4200
               TabIndex        =   41
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label4 
               Caption         =   "환자명 :"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3150
               TabIndex        =   40
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Left            =   1605
               TabIndex        =   39
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label2 
               Caption         =   "바코드번호 :"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.CommandButton cmdRSch 
            Caption         =   "로컬결과조회"
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
            Left            =   2880
            TabIndex        =   23
            Top             =   300
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   105
            TabIndex        =   22
            Top             =   300
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
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   645
            TabIndex        =   20
            Top             =   780
            Width           =   225
         End
         Begin VB.CommandButton cmdRTrans 
            Caption         =   "결과수동전송"
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
            Left            =   5505
            TabIndex        =   19
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdRClear 
            Caption         =   "화면초기화"
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
            Left            =   6990
            TabIndex        =   18
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   7815
            Left            =   105
            TabIndex        =   21
            Top             =   720
            Width           =   8280
            _Version        =   393216
            _ExtentX        =   14605
            _ExtentY        =   13785
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   13
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":A90C
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   7305
            Left            =   8460
            TabIndex        =   45
            Top             =   1260
            Width           =   5985
            _Version        =   393216
            _ExtentX        =   10557
            _ExtentY        =   12885
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":B354
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8775
         Left            =   135
         TabIndex        =   7
         Top             =   360
         Width           =   14625
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Frame Frame6 
            Height          =   585
            Left            =   8505
            TabIndex        =   55
            Top             =   630
            Width           =   6015
            Begin VB.Label Label10 
               Caption         =   "바코드번호 :"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   60
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblBarcodeNow 
               Caption         =   "1234567890ab"
               Height          =   165
               Left            =   1605
               TabIndex        =   59
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label8 
               Caption         =   "환자명 :"
               BeginProperty Font 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3150
               TabIndex        =   58
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblPnameNow 
               Caption         =   "1234567890ab"
               Height          =   225
               Left            =   4200
               TabIndex        =   57
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   56
               Top             =   720
               Width           =   1155
            End
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Command16"
            Height          =   435
            Left            =   6060
            TabIndex        =   12
            Top             =   4920
            Width           =   1215
         End
         Begin VB.TextBox txtTest 
            Height          =   675
            Left            =   1740
            TabIndex        =   11
            Top             =   4740
            Width           =   4125
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "화면초기화"
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
            Left            =   6990
            TabIndex        =   16
            Top             =   210
            Width           =   1395
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "결과수동전송"
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
            Left            =   5520
            TabIndex        =   15
            Top             =   210
            Width           =   1395
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   660
            TabIndex        =   10
            Top             =   780
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   7815
            Left            =   135
            TabIndex        =   14
            Top             =   720
            Width           =   8235
            _Version        =   393216
            _ExtentX        =   14526
            _ExtentY        =   13785
            _StockProps     =   64
            AllowMultiBlocks=   -1  'True
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   14
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            ScrollBars      =   2
            SelectBlockOptions=   8
            SpreadDesigner  =   "frmInterface.frx":F17F
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   7260
            Left            =   8505
            TabIndex        =   13
            Top             =   1275
            Width           =   5985
            _Version        =   393216
            _ExtentX        =   10557
            _ExtentY        =   12806
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":FBFC
         End
         Begin VB.Frame Frame2 
            Caption         =   "Error Log"
            Height          =   1815
            Left            =   8505
            TabIndex        =   8
            Top             =   6720
            Visible         =   0   'False
            Width           =   5970
            Begin VB.TextBox txtErrLog 
               Appearance      =   0  '평면
               Height          =   1455
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   9
               Top             =   240
               Width           =   5775
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10095
      Width           =   15045
      _ExtentX        =   26538
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
            TextSave        =   "2011-12-21"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 6:37"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Service Center (02)6205-1751"
            TextSave        =   "Service Center (02)6205-1751"
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   645
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   1138
      _Version        =   131072
      ForeColor       =   16777215
      BackColor       =   11494691
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "     DxC  INTERFACE"
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         Picture         =   "frmInterface.frx":13A4B
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   195
         Visible         =   0   'False
         Width           =   285
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   12120
         TabIndex        =   2
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   11190
         TabIndex        =   5
         Top             =   270
         Width           =   780
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
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
         Left            =   5370
         TabIndex        =   4
         Top             =   255
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin FPSpread.vaSpread vasOrder_Signal 
      Height          =   8925
      Left            =   15180
      TabIndex        =   64
      Top             =   1020
      Width           =   4725
      _Version        =   393216
      _ExtentX        =   8334
      _ExtentY        =   15743
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
      SpreadDesigner  =   "frmInterface.frx":13FD5
   End
   Begin VB.Menu MnMain 
      Caption         =   "메인"
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "설정"
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "전송"
      Begin VB.Menu MnTransAuto 
         Caption         =   "자동"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "수동"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'vasid, vasrid colum
'Const colCheckBox = 1
'Const colBarcode = 2
'Const colRack = 3
'Const colPos = 4
'Const colPID = 5
'Const colPName = 6
'Const colSex = 7
'Const colAge = 8
'Const colJumin = 9
'Const colOCnt = 10
'Const colHospital = 11
'Const colState = 12


Const colCheckBox = 1
Const colSpecNo = 2
Const colBarcode = 3
Const colRack = 4
Const colPos = 5
Const colPID = 6
Const colPName = 7
Const colSex = 8
Const colAge = 9
Const colOCnt = 10
Const colRCnt = 11
Const colState = 12
Const colTestType = 13
'Const colA1c = 13
'Const colIFCC = 15
'Const coleAg = 17


'sendflag
'0: Order
'1: Result
'2: Trans

'vasres, vasrres colum
Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
Const colResult = 4
Const colSeq = 5
Const colFLAG = 6
Const colEquipResult = 7
Const colDelta = 8
Const colPanic = 9

Dim gRow As Long
Dim gsBarCode As String
Dim gsSampleType As String
Dim gsPID As String
Dim gsRackNo As String
Dim gsPosNo As String
Dim gsResDateTime As String
Dim gsSeqNo As String
Dim gsExamCode As String
Dim gsExamName As String
Dim gsOrder As String
Dim gsResult As String
Dim gsFlag As String
Dim gsTestID As String

Dim gMT As String
Dim gComState As Long
Dim gErrState As Long

Dim gIFCC1 As String
Dim gIFCC2 As String
Dim geAg1 As String
Dim geAg2 As String
Dim gADD_IFCC As String
Dim gADD_eAg As String


Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
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

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow + 1
    vasActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
End Sub

Private Sub cmdOrder_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
    
Private Sub cmdQC_Click()
    'frmQCResSch.Show
End Sub

Private Sub cmdResCall_Click()
'    frmResult.Show 0
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    lRow = vasID.ActiveRow
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow - 1
    vasActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1
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

Private Sub Command14_Click()
'    frmUserChange.Show 0
    
End Sub

Private Sub chkRAll_Click()
    Dim iRow As Long
    
    If chkRAll.Value = 1 Then
        For iRow = 1 To vasRID.DataRowCnt
            vasRID.Row = iRow
            vasRID.Col = 1
            
            vasRID.Value = 1
        Next iRow
    ElseIf chkRAll.Value = 0 Then
        For iRow = 1 To vasRID.DataRowCnt
            vasRID.Row = iRow
            vasRID.Col = 1
            
            vasRID.Value = 0
        Next iRow
    End If
End Sub

''Private Sub cmdExcel_Click()
''    Dim iRow As Integer
''    Dim j As Integer
''
''    Dim sCurDate As String
''    Dim sSerDate As String
''    Dim sHead As String
''    Dim sFoot As String
''    Dim sFileName As String
''
''    Dim sA1c As String
''    Dim sIFCC As String
''    Dim seAg As String
''
''
''
''    ClearSpread vasPrint
''
''    j = 1
''
''    For iRow = 1 To vasRID.DataRowCnt
''        vasRID.Row = iRow
''        vasRID.Col = 1
''
''        If vasRID.Value = 1 Then
''            SetText vasPrint, Trim(GetText(vasRID, iRow, colSpecNo)), j, 1
''            SetText vasPrint, Trim(GetText(vasRID, iRow, colBarcode)), j, 2
''            SetText vasPrint, Trim(GetText(vasRID, iRow, colPID)), j, 3
''            SetText vasPrint, Trim(GetText(vasRID, iRow, colPName)), j, 4
''            SetText vasPrint, Trim(GetText(vasRID, iRow, colSex)), j, 5
''            'SetText vasPrint, Trim(GetText(vasRID, iRow, colHospital)), j, 5
''
''            SQL = "SELECT RESULT " & vbCrLf & _
''                  "FROM PAT_RES " & vbCrLf & _
''                  "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
''                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
''                  "  AND BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf & _
''                  "  AND PID = '" & Trim(GetText(vasPrint, iRow, 3)) & "' " & vbCrLf & _
''                  "ORDER BY SEQNO"
''            res = db_select_Vas(gLocal, SQL, vasPrintBuf)
''
''            sA1c = GetText(vasPrintBuf, 1, 1)
''            sIFCC = GetText(vasPrintBuf, 2, 1)
''            seAg = GetText(vasPrintBuf, 3, 1)
''
''            ClearSpread vasPrintBuf, 1, 1
''
''            SetText vasPrint, sA1c, j, 7
''            SetText vasPrint, sIFCC, j, 8
''            SetText vasPrint, seAg, j, 9
''
''            '"GROUP BY BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, JUMIN, Hospital, SENDFLAG"
''
'''            SetText vasprint, Trim(GetText(vasrid, iRow, vasrid.MaxCols)), j, 8
'''            SetText vasprint, Trim(GetText(vasrid, iRow, 10)), j, 9
''
''            j = j + 1
''        End If
''    Next iRow
''
''    If vasPrint.DataRowCnt < 1 Then
''        MsgBox "저장할 자료가 없습니다.", , "알 림"
''        Exit Sub
''    Else
''        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
''        CommonDialog1.ShowSave
''        sFileName = CommonDialog1.Filename
''        SaveExcel sFileName, vasPrint
''
''    End If
''End Sub
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

Private Sub cmdIFClear_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
'    dtptoday = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            
            If Mid(GetText(vasID, lRow, colBarcode), 1, 2) = "99" Then
                res = Insert_Data_QC(lRow)
            Else
                 res = Insert_Data(lRow)
            End If
            
            
            If res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasID, lRow, colBarcode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasID.Row = lRow
            vasID.Col = 1
            vasID.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdRClear_Click()
    Dim i As Integer

'    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasRID, 1, vasRID.MaxRows, 1, vasRID.MaxCols, 0, 0, 0
    SetForeColor vasRRes, 1, vasRRes.MaxRows, 1, vasRRes.MaxCols, 0, 0, 0
    
    vasRID.MaxRows = 0
    vasRRes.MaxRows = 0
    
    dtpExamDate = Date
    
End Sub

Private Sub cmdRSch_Click()
    Dim iRow As Long

    ClearSpread vasRID
    ClearSpread vasRRes
    Call chkRAll_Click
    
    SQL = "SELECT '', RECENO, BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, COUNT(*), COUNT(*), SENDFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND SENDFLAG IN ('1', '2','0') " & vbCrLf & _
          "GROUP BY BARCODE, RECENO, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, SENDFLAG"
    res = db_select_Vas(gLocal, SQL, vasRID)
    
          '"  AND SENDFLAG IN ('1', '2') "
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For iRow = 1 To vasRID.DataRowCnt
        Select Case Trim(GetText(vasRID, iRow, colState))
        Case "2"
            SetBackColor vasRID, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasRID, "완료", iRow, colState
        Case "0"
            SetText vasRID, "오더", iRow, colState
        Case "1"
            SetText vasRID, "결과", iRow, colState
        End Select
    Next iRow
End Sub

Private Sub cmdRTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasRID.DataRowCnt
        vasRID.Row = lRow
        vasRID.Col = 1
        If vasRID.Value = 1 Then
            
            If Mid(GetText(vasRID, lRow, colBarcode), 1, 2) = "99" Then
                res = Insert_Data_QC_R(lRow)
            Else
                res = Insert_Data_R(lRow)
            End If
        
            If res = -1 Then
                SetForeColor vasRID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasRID, "Failed", lRow, colState
            ElseIf res = 0 Then
            
            Else
                vasRID.Row = lRow
                vasRID.Col = 1
                vasRID.Value = 1
                
                SetBackColor vasRID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasRID, "Trans", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasRID, lRow, colBarcode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasRID.Row = lRow
            vasRID.Col = 1
            vasRID.Value = 0
        End If
    Next lRow
End Sub

Private Sub Command1_Click()
    Dim iRow As Long

    ClearSpread vasRID
    ClearSpread vasRRes
    Call chkAll_Click
    
    SQL = "SELECT '', RECENO, BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, COUNT(*), COUNT(*), SENDFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND SENDFLAG IN ('1', '2') " & vbCrLf & _
          "GROUP BY BARCODE, RECENO, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, SENDFLAG"
    res = db_select_Vas(gLocal, SQL, vasID)
    
          '"  AND SENDFLAG IN ('1', '2') "
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For iRow = 1 To vasID.DataRowCnt
        Select Case Trim(GetText(vasID, iRow, colState))
        Case "2"
            SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasRID, "완료", iRow, colState
'        Case "0"
'            SetText vasID, "오더", iRow, colState
        Case "1"
            SetText vasID, "결과", iRow, colState
        End Select
    Next iRow
End Sub

Private Sub Form_Activate()
    Call GetExamCode
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape And fraResFlag.Visible = True Then
        fraResFlag.Visible = False
    End If
End Sub

Private Sub lblclear_Click()
    lblPnameNow.Caption = ""
    lblBarcodeNow.Caption = ""
    lblChangeBar.Caption = ""
    lblBarcode.Caption = ""
    lblChangePID.Caption = ""
    lblPname.Caption = ""
End Sub

Private Sub Command16_Click()
    Dim lsChar As String
    Dim i As Integer
    Dim sSendData As String
    Dim CharData As String
    
    For i = 1 To Len(txtTest)
    
    
    lsChar = Mid(txtTest, i, 1)
    
    Select Case lsChar
    Case chrSOH
        txtData.Text = txtData.Text & lsChar
        Save_Raw_Data "[Rx]" & lsChar
        gPreMsg = chrACK
        MSComm1.Output = chrACK
        Save_Raw_Data "[Tx]" & chrACK
        gACKSig = 1
        gComState = 0
        
    Case "["
        txtData.Text = lsChar
        
    Case chrLF
        txtData.Text = txtData.Text & lsChar
        
        Save_Raw_Data "[Rx]" & txtData.Text
                
         LX20 Mid(txtData.Text, 2)
        gComState = 1
        
        If gACKSig = 1 Then
            gPreMsg = chrETX
            gACKSig = 0
        Else
            gPreMsg = chrACK
            gACKSig = 1
        End If
        MSComm1.Output = gPreMsg
        Save_Raw_Data "[Tx]" & gPreMsg
        
        txtData = ""
    Case chrEOT
        txtData.Text = lsChar
        
        If gComState = 1 And vasTemp1.DataRowCnt > 0 Then
            gPreMsg = chrEOT & chrSOH
            MSComm1.Output = chrEOT & chrSOH
            Save_Raw_Data "[Tx]" & chrEOT & chrSOH
            
            gComState = 2
        End If
    Case chrACK
        Save_Raw_Data "[Rx]" & chrACK
        
        'If gComState = 2 Then
            gOrderMessage = GetText(vasOrder_Signal, 1, 1)
            vasOrder_Signal.DeleteRows 1, 1
            gPreMsg = gOrderMessage
            MSComm1.Output = gOrderMessage
            Save_Raw_Data "[Tx]" & gOrderMessage
            gOrderMessage = ""
            gComState = 3
'        ElseIf gComState = -1 Then
'            CX_Init
        'End If
    Case chrETX
        Save_Raw_Data "[Rx]" & chrACK
        
        gPreMsg = chrEOT
        MSComm1.Output = chrEOT
        Save_Raw_Data "[Tx]" & chrEOT
        
        If vasTemp1.DataRowCnt > 0 Then
            gPreMsg = chrEOT & chrSOH
            MSComm1.Output = chrEOT & chrSOH
            Save_Raw_Data "[Tx]" & chrEOT & chrSOH
            
            gComState = 2
        Else
            gComState = 0
        End If
    
    Case Else
        txtData.Text = txtData.Text & lsChar
    End Select
    Next i

    txtTest = ""
End Sub

Sub URISCAN_PRO(asData As String)
    Dim MyVar As String
    Dim MyRet As String
          
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim iRow As Integer
    Dim jRow As Integer
    Dim llRow As Integer
    Dim liRet As Integer
    
    Dim sBarcode As String
    Dim sEquipCode As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sSeqNo As String
    Dim sResult As String
    
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sDate As String
    
    Dim lsSeq As String
    Dim lsCnt As String
    
    If Trim(asData) = "" Then
        Exit Sub
    End If
    
    MyVar = Trim(asData)
         
    sDate = Format(dtpToday, "yyyymmdd")
    
    i = InStr(1, MyVar, "Date")
    If i > 0 Then
        sDate = Format(CDate(Trim(Mid(MyVar, i + 6, 20))), "yyyy-mm-dd hh:nn:ss")
    End If
    
    i = InStr(1, MyVar, "ID_NO")
    sSeqNo = CStr(CLng(Trim(Mid(MyVar, i + 6, 4))))

    sBarcode = CStr(Trim(Mid(MyVar, i + 11, 12)))
    
    '같은 바코드번호의 검체는 디스플레이되지 않음
    llRow = -1
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, 0)) = sSeqNo Then
            llRow = iRow
            Exit For
        End If
        
        If Trim(GetText(vasID, iRow, 0)) = "" Then
            llRow = iRow
            Exit For
        End If
    Next iRow

    If llRow = -1 Then
        llRow = vasID.DataRowCnt + 1
        If llRow > vasID.MaxRows Then
            vasID.MaxRows = llRow
        End If
    End If
    
    ClearSpread vasRes, 1, 1

    SetText vasID, sSeqNo, llRow, 0
    'SetText vasID, sExamDate, llRow, colDate
    'SetText vasID, sDate, llRow, colTime
    SetText vasID, sBarcode, llRow, colBarcode
    
    '수신중========================================================
    SetText vasID, "수신중", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205
    '==============================================================
    
    '샘플의 환자 정보 가져오기
    Get_Sample_Info llRow
    
    '검사코드만큼 Row의 갯수를 설정
    gReadBuf(0) = "0"
    
    SQL = "Select count(examcode) From equipexam" & vbCrLf & _
          " Where equipno = '" & gEquip & "' "
    res = db_select_Col(gLocal, SQL)

    vasRes.MaxRows = Trim(gReadBuf(0))

    
    lsSeq = ""
    lsCnt = ""
        
    
    '결과 잘라 넣기
    j = 0
    For j = 1 To vasRes.MaxRows
        sExamName = Trim(GetText(vasCode, j, 1))
        
        Select Case sExamName
        Case "BLD", "BIL", "URO", "KET", "PRO", "NIT", "GLU", "LEU"
            i = InStr(1, MyVar, Trim(sExamName))
            sResult = Trim(Mid(MyVar, i + 3, 8))

        Case "p.H"
            i = InStr(1, MyVar, "p.H")
            sResult = Trim(Mid(MyVar, i + 3, 14))

        Case "S.G"
            i = InStr(1, MyVar, "S.G")

            If Mid(MyVar, i) = "<=" Or Mid(MyVar, i) = ">=" Then
                sResult = Trim(Mid(MyVar, i + 3, 9))
            Else
                sResult = Trim(Mid(MyVar, i + 3, 12))
            End If
        End Select
        
        Select Case sResult
        Case "-"
            sResult = "Negatvie"
        End Select
        
        ClearSpread vasTemp
        
        SQL = "Select examcode, '', examname From EquipExam" & vbCrLf & _
              " Where Equipno = '" & gEquip & "' " & vbCrLf & _
              "  And EquipCode = '" & Trim(sExamName) & "'"
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
        For jRow = 1 To vasTemp.DataRowCnt
            sExamCode = Trim(GetText(vasTemp, jRow, 1))
            sSeqNo = Trim(GetText(vasTemp, jRow, 2))
            sExamName = Trim(GetText(vasTemp, jRow, 3))
        
            SetText vasRes, Trim(sExamName), j, colEquipCode '장비코드
            SetText vasRes, sExamCode, j, colExamCode '검사코드
            SetText vasRes, sExamName, j, colExamName '검사명
            SetText vasRes, Trim(sResult), j, colResult   '검사결과
            SetText vasRes, sSeqNo, j, colSeq        '순번(서브코드)
            Trim (GetText(vasID, llRow, 0))
            Save_Local_One llRow, j, "1", CStr(Trim(sResult))
        Next jRow
    Next j
    gReadBuf(0) = ""
    
    '수신중========================================================
    SetText vasID, "수신완료", llRow, colState
    SetBackColor vasID, llRow, llRow, 1, 1, 0, 128, 64
    '==============================================================
    

End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    If App.PrevInstance Then
        End
    End If

    cmdIFClear_Click
    cmdRClear_Click
    lblclear_Click

    GetSetup
    
    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
     '-- osw 추가
    For i = 1 To 3
        If Not Connect_PRServer Then
            cn_cnt = cn_cnt + 1
            If cn_cnt = 3 Then
                If Not Connect_DRServer Then
                    MsgBox "연결되지 않았습니다."
                    cn_Server_Flag = False
                    Exit Sub
                Else
                    cn_Server_Flag = True
                End If
            End If
        Else
            cn_Server_Flag = True
        End If
    Next

    GetExamCode
    dtpToday = Date
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -30), "yyyymmdd")
    
    SQL = "delete from pat_res where examdate < '" & sDate & "'"
    res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
    If lblUser.Caption = "" Then
        Call picLogin_Click
    End If
    stInterface.Tab = 0
End Sub



Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  examcode "
    res = db_select_Vas(gLocal, SQL, vasCode)
    If res > 0 Then
        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 6)
    Else
        SaveQuery SQL
        Exit Function
    End If
        
    For i = 1 To vasCode.DataRowCnt
        If i = 1 Then
            gAllExam = "'" & Trim(GetText(vasCode, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ",'" & Trim(GetText(vasCode, i, 2)) & "'"
        End If
        
        gArrEquip(i, 1) = i
        For j = 1 To 5
            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
        Next j
        
        
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Server
    DisConnect_Local
    Unload Me
    End
End Sub

Private Sub MnExamConfig_Click()
    frmOrderCode.Show
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show
End Sub

Private Sub MnTransAuto_Click()
    chkMode.Caption = "Auto"
    MnTransAuto.Checked = True
    MnTransManual.Checked = False
    chkMode.Value = 1
    
End Sub

Private Sub MnTransManual_Click()
    chkMode.Caption = "Manual"
    MnTransAuto.Checked = False
    MnTransManual.Checked = True
    chkMode.Value = 0
End Sub

Private Sub MSComm1_OnComm()
        Dim lsChar As String
    
    lsChar = MSComm1.Input
    
    Select Case lsChar
    Case chrSOH
        txtData.Text = txtData.Text & lsChar
        gPreMsg = chrACK
        MSComm1.Output = chrACK
        Save_Raw_Data "[Tx]" & chrACK
        gACKSig = 1
        gComState = 0
        
    Case "["
        txtData.Text = lsChar
        
    Case chrLF
        txtData.Text = txtData.Text & lsChar
        
        Save_Raw_Data "[Rx]" & txtData.Text
                
        LX20 Mid(txtData.Text, 2)
        gComState = 1
        
        If gACKSig = 1 Then
            gPreMsg = chrETX
            gACKSig = 0
        Else
            gPreMsg = chrACK
            gACKSig = 1
        End If
        MSComm1.Output = gPreMsg
        Save_Raw_Data "[Tx]" & gPreMsg
        
        txtData = ""
    Case chrEOT
        txtData.Text = lsChar
        
        If gComState = 1 And vasOrder_Signal.DataRowCnt > 0 Then
            gPreMsg = chrEOT & chrSOH
            MSComm1.Output = chrEOT & chrSOH
            Save_Raw_Data "[Tx]" & chrEOT & chrSOH
            
            gComState = 2
        End If
    Case chrACK
        Save_Raw_Data "[Rx]" & chrACK
        
        If gComState = 2 Then
            gOrderMessage = GetText(vasOrder_Signal, 1, 1)
            vasOrder_Signal.DeleteRows 1, 1
            gPreMsg = gOrderMessage
            MSComm1.Output = gOrderMessage
            Save_Raw_Data "[Tx]" & gOrderMessage
            gOrderMessage = ""
            gComState = 3
'        ElseIf gComState = -1 Then
'            CX_Init
        End If
    Case chrETX
        Save_Raw_Data "[Rx]" & chrACK
        
        gPreMsg = chrEOT
        MSComm1.Output = chrEOT
        Save_Raw_Data "[Tx]" & chrEOT
        
        If vasOrder_Signal.DataRowCnt > 0 Then
            gPreMsg = chrEOT & chrSOH
            MSComm1.Output = chrEOT & chrSOH
            Save_Raw_Data "[Tx]" & chrEOT & chrSOH
            
            gComState = 2
        Else
            gComState = 0
        End If
    
    Case Else
        txtData.Text = txtData.Text & lsChar
    End Select

End Sub

Sub LX20(asVar As String)
    Dim lsBuff As String
    Dim lsData As String
    Dim i, j, k, X As Integer
    Dim iField As Integer
    Dim lsStream As String
    Dim lsFunction As String
    Dim ii As Integer
    
    
    Dim lRow As Long
    
    Dim lsID As String
    Dim lsDisk  As String
    Dim lsPos   As String
    Dim lsEquipCode As String
    Dim lsExamCode As String
    Dim lsResult As String
    Dim lsResult1 As String
    Dim lsResFlag As String
    Dim lsTestType As String
    
    Dim Result_Date As String
    Dim Result_Time As String
    Dim Result_Date_Time As String
    
    '//////// HL , Panic, Delta 체크
    Dim Ref_Flag        As String
    Dim Ref_Cnt         As Integer
    Dim Ref_HL          As String
    Dim Ref_Delta       As String
    Dim Ref_Panic       As String
    
    '//////// 검사 결과
    Dim sResType As String
        
    gOrderMessage = ""
    
    i = InStr(1, asVar, "]")
    If i > 0 Then
        lsBuff = Left(asVar, i - 1) & ","
    Else
        lsBuff = asVar
    End If
    
    i = InStr(1, lsBuff, ",")
    lsData = Left(lsBuff, i - 1)
    lsBuff = Mid(lsBuff, i + 1)
    i = InStr(1, lsBuff, ",")
    lsData = Left(lsBuff, i - 1)
    lsStream = Trim(lsData)
    lsBuff = Mid(lsBuff, i + 1)
    i = InStr(1, lsBuff, ",")
    lsData = Left(lsBuff, i - 1)
    lsFunction = Trim(lsData)
    lsBuff = Mid(lsBuff, i + 1)
    
   If Trim(lsStream) = "801" And Trim(lsFunction) = "06" Then   'Host Query
        Call vasOrderBuf.DeleteRows(1, 500)
   
        i = InStr(1, lsBuff, ",")
'        i = InStr(1, lsBuff, "    ")
        Do While i > 0
            lsData = Left(lsBuff, i - 1)
            lsID = Trim(lsData)
            
            If Trim(lsID) <> "" And IsNumeric(lsID) = True And Mid(lsID, 1, 2) <> "99" Then
                Call Proc_Order_LX(lsID)
                If dtpToday <> Format(Date, "yyyy/mm/dd") Then
                    dtpToday = Format(Date, "yyyy/mm/dd")
                End If
            ElseIf Trim(lsID) <> "" And (IsNumeric(lsID) = False Or Mid(lsID, 1, 2) = "99") Then
                Call Proc_Order_LX_QC(lsID)
                If dtpToday <> Format(Date, "yyyy/mm/dd") Then
                dtpToday = Format(Date, "yyyy/mm/dd")
                End If
            End If
            
            lsBuff = Mid(lsBuff, i + 1)
'            i = InStr(1, lsBuff, "    ")
            i = InStr(1, lsBuff, ",")
        Loop
        
        
    ElseIf Trim(lsStream) = "802" And Trim(lsFunction) = "01" Then   'Results : Cup Head
        iField = 3
        i = InStr(1, lsBuff, ",")
        Do While i > 0
            iField = iField + 1
            
            lsData = Left(lsBuff, i - 1)
    
            Select Case iField
            Case 1  'Device ID
            Case 2  'Stream
            Case 3  'Function
            Case 4  'Date Start
                Result_Date = Trim(lsData)
                If dtpToday <> Format(Date, "yyyy/mm/dd") Then
                    dtpToday = Format(Date, "yyyy/mm/dd")
                End If
                '12/01/2010 00:16:09
                'Result_Date = Mid(Result_Date, 1, 2) & "/" & Mid(Result_Date, 3, 2) & "/" & Mid(Result_Date, 5)
            Case 5  'Time Start
                'Result_Time = Trim(lsData)
                'Result_Time = Mid(Result_Time, 3, 2) & ":" & Mid(Result_Time, 1, 2) & ":" & Mid(Result_Time, 5)
                
                'Result_Date_Time = Result_Date & " " & Result_Time
            Case 6  'Accession Number
            Case 7  'Print Type
            Case 8  'Sector Number
                lsDisk = Trim(lsData)
            Case 9  'Cup Number
                lsPos = Trim(lsData)
            Case 10 'Test Type
            Case 11 'Future Use Space
            Case 12 'SampleType
                lsTestType = Trim(lsData)
                
            Case 13 'Sample ID
                
                lsID = Trim(lsData)
                gRow = -1
                ClearSpread vasResTemp
                ClearSpread vasRes
                For lRow = 1 To vasID.DataRowCnt
                    If Trim(GetText(vasID, lRow, colBarcode)) = lsID And Trim(GetText(vasID, lRow, colRack)) = "" Then
                        gRow = lRow
                        Exit For
                    End If
                Next lRow
                If gRow < 1 Then
                    gRow = vasID.DataRowCnt + 1
                    If vasID.MaxRows < gRow Then vasID.MaxRows = gRow
                End If
                
                SetText vasID, lsID, gRow, colBarcode
                SetText vasID, lsDisk, gRow, colRack
                SetText vasID, lsPos, gRow, colPos
                SetText vasID, lsTestType, gRow, colTestType
                
                '환자정보 불러오기
                If Len(lsID) = 10 And Mid(lsID, 1, 2) = "99" Then   '2010.08.16 이상은
                    Get_Sample_Info_QC gRow
                        
                    '///////////////// QC Rerun땜에 넣음
                    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                        "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
                        "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                        "  AND BARCODE = '" & Trim(GetText(vasID, lRow, colBarcode)) & "' "
                    res = SendQuery(gLocal, SQL)
                    
                    ClearSpread vasList
                ElseIf Len(lsID) = 10 And Mid(lsID, 1, 2) <> "99" Then
                    Get_Sample_Info gRow
                    ClearSpread vasList
                    Call Pat_Info(lsID)
                End If
                
                '////// 결과 저장용 버퍼에도 넣어줌
                SetText vasList, GetText(vasID, gRow, colSpecNo), 1, colSpecNo
                SetText vasList, lsID, 1, colBarcode
                SetText vasList, lsDisk, 1, colRack
                SetText vasList, lsPos, 1, colPos
                SetText vasList, GetText(vasID, gRow, colPID), 1, colPID
                SetText vasList, GetText(vasID, gRow, colPName), 1, colPName
                SetText vasList, GetText(vasID, gRow, colAge), 1, colAge
                SetText vasList, GetText(vasID, gRow, colSex), 1, colSex
                SetText vasList, Result_Date_Time, 1, colTestType + 1
                
                Save_Local_One_Signal1 1, i, "1", vasResTemp, lsResult1
                
                Exit Do
            End Select
    
            lsBuff = Mid(lsBuff, i + 1)
            i = InStr(1, lsBuff, ",")
        Loop
        
        
    ElseIf Trim(lsStream) = "802" And Trim(lsFunction) = "03" Then   'Test Results
        iField = 3
        i = InStr(1, lsBuff, ",")
        Do While i > 0
            iField = iField + 1
            
            lsData = Left(lsBuff, i - 1)
    
            Select Case iField
            Case 1  'Device ID
            Case 2  'Stream
            Case 3  'Function
            Case 4  'Date Complete
            Case 5  'Time Complete
            Case 6  'Accession Number
            Case 7  'Result Record Number
            Case 8  'Rack Number
                lsDisk = Trim(lsData)
            Case 9  'Cup Number
                lsPos = Trim(lsData)
            Case 10 'Sample ID
                lsID = Trim(lsData)
            Case 11 'Chemistry
                lsEquipCode = Trim(lsData)
            Case 12 'Reagent Serial No.
            Case 13 'Reagent Lot No.
            Case 14 'Cuvette No.
            Case 15 'Replicate No.
            Case 16 'Result in Selected Units
                lsResult = Trim(lsData)
                lsResult1 = lsResult
                If Left(lsResult, 1) = "#" Or Left(lsResult, 1) = "*" Then
                    lsResult = ""
                End If

            Case 17 'Calibration Rate
            Case 18 'Positive or Negative
            Case 19 'Suppress Result
            Case 20 'Units
            Case 21 'Normal Range Flag
            Case 22 'Instrument Range Flag
            Case 27 'Err Remark
                lsResFlag = Trim(lsData)
                If lsResult = "" Then: lsResult = "##" & lsResFlag
                '2010.08.16 이상은 - QC결과 BNP 관련 처리
                If Len(lsID) <> 10 And Trim(GetText(vasID, gRow, colTestType)) = "PL" Then
                    lsTestType = "SE"
                    SetText vasID, lsTestType, gRow, colTestType
                End If
                'lsResult1 = lsResult    '///// 장비결과(1)
                ClearSpread vasTemp
                
                If Mid(lsID, 1, 2) = "99" Then
                    Call EquipExamCode_QC(lsEquipCode, lsID)
                Else
                    If lsEquipCode = "03E" Then Call eGFR_SAVE(lsID, lsResult)
                    Call EquipExamCode(lsEquipCode, lsID)
                    
                    If EXAMCODE_LIMIT(gEquipExamCode, lsResult) <> "" Then
                        lsResult = EXAMCODE_LIMIT(gEquipExamCode, lsResult)
                    End If
                    
                    If gEquipExamCode = "L3537" Then
                        If lsEquipCode = "89A" Then
                            If IsNumeric(lsResult1) = True Then
                                If lsResult1 < 1 Then
                                    lsResult = "Negative"
                                Else
                                    lsResult = "Positive(" & lsResult1 & ")"
                                End If
                            Else
                                'lsResult1 = lsResult
                                Do
                                    lsResult = Mid(lsResult, 2)
                                    If IsNumeric(Mid(lsResult, 1, 1)) = True Then
                                        If InStr(1, lsResult, ")") > 0 Then: lsResult = Mid(lsResult, 1, InStr(1, lsResult, ")") - 1)
                                        Exit Do
                                    End If
                                    If Len(lsResult) = 1 Then lsResult = lsResult1: Exit Do
                                Loop
                                If IsNumeric(lsResult) = True Then
                                    If lsResult < 1 Then
                                        lsResult = "Negative"
                                    Else
                                        lsResult = "Positive(" & lsResult1 & ")"
                                    End If
                                End If
                            End If
                            
                        End If
                    End If
                End If
                
                
                If IsNumeric(gExamRange) = True Then
                    For k = 0 To gExamRange
                                If k = 0 Then
                            sResType = "#0"
                        ElseIf k = 1 Then
                            sResType = sResType & ".0"
                        Else
                            sResType = sResType & "0"
                        End If
                    Next
                    If IsNumeric(lsResult) = True Then lsResult = Format(lsResult, sResType)
                Else
                
                End If
                
                SQL = "Select examcode, examname, reflow, refhigh, resprec from equipexam" & vbCrLf & _
                      "Where equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and examcode  = '" & gEquipExamCode & "' " & vbCrLf & _
                      "  And equipcode = '" & lsEquipCode & "' "
                res = db_select_Col(gLocal, SQL)                        '///// 검사 결과 저장용
                
                '////// HL, Delta, Panuic  체크
                Ref_Flag = GetDecision(gRow, lsID, Trim(gReadBuf(0)), lsResult)
                Ref_Cnt = InStr(1, Ref_Flag, "/")
                Ref_HL = Mid(Ref_Flag, 1, Ref_Cnt - 1)
                
                Ref_Flag = Mid(Ref_Flag, Ref_Cnt + 1)
                Ref_Cnt = InStr(1, Ref_Flag, "/")
                Ref_Delta = Mid(Ref_Flag, 1, Ref_Cnt - 1)
                
                Ref_Flag = Mid(Ref_Flag, Ref_Cnt + 1)
                Ref_Panic = Ref_Flag
                
                
                '///// 결과 입력
                i = vasResTemp.DataRowCnt + 1
                
                SetText vasResTemp, lsEquipCode, i, colEquipCode     '장비코드
                SetText vasResTemp, Trim(gReadBuf(0)), i, colExamCode    '검사코드
                SetText vasResTemp, Trim(gReadBuf(1)), i, colExamName    '검사명
                SetText vasResTemp, lsResult, i, colResult              '검사결과
                SetText vasResTemp, lsResult1, i, colEquipResult            '장비 검사결과
                
                SetText vasResTemp, Ref_HL, i, colFLAG
                SetText vasResTemp, Ref_Delta, i, colDelta
                SetText vasResTemp, Ref_Panic, i, colPanic
                
                Save_Local_One_Signal 1, i, "1", vasResTemp, lsResult1

                Exit Do
            Case 23 'Critical Range Flag
            Case 24 'ORDAC Result
            Case 25 'Control Range Flag
            Case 26 'Calculated Result
'                lsResult1 = Trim(lsData)
'                If Left(lsResult1, 1) = "#" Or Left(lsResult1, 1) = "*" Then
'                    lsResult1 = ""
'                End If
            Case 27 'Instruments Code
            End Select
    
            lsBuff = Mid(lsBuff, i + 1)
            i = InStr(1, lsBuff, ",")
        Loop
        
        
    ElseIf Trim(lsStream) = "802" And Trim(lsFunction) = "11" Then   'Results : Cup Head
        iField = 3
        i = InStr(1, lsBuff, ",")
        Do While i > 0
            iField = iField + 1
            
            lsData = Left(lsBuff, i - 1)
    
            Select Case iField
            Case 1  'Device ID
            Case 2  'Stream
            Case 3  'Function
            Case 4  'Date Start
            Case 5  'Time Start
            Case 6  'Accession Number
            Case 7  'Print Type
                lsDisk = Trim(lsData)
            Case 8  'Sector Number
                lsPos = Trim(lsData)
            Case 9  'Cup Number
                lsID = Trim(lsData)
                gRow = -1
                ClearSpread vasResTemp
                ClearSpread vasRes
                For lRow = 1 To vasID.DataRowCnt
                    If Trim(GetText(vasID, lRow, colBarcode)) = lsID And Trim(GetText(vasID, lRow, colRack)) = lsDisk Then
                        gRow = lRow
                        Exit For
                    End If
                Next lRow
                If gRow < 1 Then
                    gRow = vasID.DataRowCnt + 1
                    If vasID.MaxRows < gRow Then vasID.MaxRows = gRow
                End If
            Case 10 'Test Type
            Case 11 '장비코드
                lsEquipCode = Trim(lsData)
            Case 12 '에러 체크
                lsResFlag = Trim(lsData)
                
                '2010.08.09 이상은 - PL검체(L61)인데 SE(L60)로 해서 결과전송이 안 되는지?
'                If lsTestType = "PL" Then
'                    lsTestType = "SE"
'                End If
                '****************************************************
            Case 13 '계산 오더 결과
                
                lsResult = Trim(lsData)
                
                If Left(lsResult, 1) = "#" Or Left(lsResult, 1) = "*" Then
                    lsResult = ""
                End If
                
                If lsEquipCode = "89A" Then
                    If IsNumeric(lsResult) = True Then
                        If lsResult < 1 Then
                            lsResult1 = "Negative"
                        Else
                            lsResult1 = "Positive(" & lsResult & ")"
                        End If
                    End If
                End If
                
                
                If lsResult = "" Then: lsResult = "##" & lsResFlag
                
                If Len(lsID) <> 10 And Trim(GetText(vasID, gRow, colTestType)) = "PL" Then
                    lsTestType = "SE"
                    SetText vasID, lsTestType, gRow, colTestType
                End If
                
                lsResult1 = lsResult    '///// 장비실제 결과 ( 수정할 내역이 있으면 추가 수정 예정)
                
                ClearSpread vasTemp
                
                If Mid(lsID, 1, 2) = "99" Then
                    Call EquipExamCode_QC(lsEquipCode, lsID)
                Else
                    
                    If lsEquipCode = "03E" Then: Call eGFR_SAVE(lsID, lsResult)
                    
                    Call EquipExamCode(lsEquipCode, lsID)
                    
                    If EXAMCODE_LIMIT(gEquipExamCode, lsResult) <> "" Then
                    lsResult = EXAMCODE_LIMIT(gEquipExamCode, lsResult)
                    End If
                    
                End If
                
                If IsNumeric(gExamRange) = True Then
                    For k = 0 To gExamRange
                                If k = 0 Then
                            sResType = "#0"
                        ElseIf k = 1 Then
                            sResType = sResType & ".0"
                        Else
                            sResType = sResType & "0"
                        End If
                    Next

                    lsResult = Format(lsResult, sResType)
                Else
                    
                    If IsNumeric(lsResult) = True Then lsResult = Format(lsResult, "#.00")
                End If
                
                SQL = "Select examcode, examname, reflow, refhigh, resprec from equipexam" & vbCrLf & _
                      "Where equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and examcode  = '" & gEquipExamCode & "' " & vbCrLf & _
                      "  And equipcode = '" & lsEquipCode & "' "
                res = db_select_Col(gLocal, SQL)                        '///// 검사 결과 저장용
                    
                '////// HL, Delta, Panuic  체크
                Ref_Flag = GetDecision(gRow, lsID, Trim(gReadBuf(0)), lsResult)
                Ref_Cnt = InStr(1, Ref_Flag, "/")
                Ref_HL = Mid(Ref_Flag, 1, Ref_Cnt - 1)
                
                Ref_Flag = Mid(Ref_Flag, Ref_Cnt + 1)
                Ref_Cnt = InStr(1, Ref_Flag, "/")
                Ref_Delta = Mid(Ref_Flag, 1, Ref_Cnt - 1)
                
                Ref_Flag = Mid(Ref_Flag, Ref_Cnt + 1)
                Ref_Panic = Ref_Flag
                
                
                '///// 결과 입력
                i = vasResTemp.DataRowCnt + 1
                
                SetText vasResTemp, lsEquipCode, i, colEquipCode     '장비코드
                SetText vasResTemp, Trim(gReadBuf(0)), i, colExamCode    '검사코드
                SetText vasResTemp, Trim(gReadBuf(1)), i, colExamName    '검사명
                SetText vasResTemp, lsResult, i, colResult              '검사결과
                SetText vasResTemp, lsResult1, i, colEquipResult            '장비 검사결과
                
                SetText vasResTemp, Ref_HL, i, colFLAG
                SetText vasResTemp, Ref_Delta, i, colDelta
                SetText vasResTemp, Ref_Panic, i, colPanic
                
                Save_Local_One_Signal 1, i, "1", vasResTemp, lsResult1
                    

                Exit Do
            End Select
    
            lsBuff = Mid(lsBuff, i + 1)
            i = InStr(1, lsBuff, ",")
        Loop
        
        
    ElseIf Trim(lsStream) = "802" And Trim(lsFunction) = "05" Then   'Results : Cup End
        If gRow < 1 Or gRow > vasID.DataRowCnt Then Exit Sub
        
        SetText vasID, "결과", gRow, colState
'        SetBackColor vasID, gRow, gRow, colBarcode, colState, 202, 255, 112
            
        If MnTransAuto.Checked = False Then Exit Sub
        
        If Len(Trim(GetText(vasID, gRow, colBarcode))) <> 10 Then Exit Sub
        
        If Mid(GetText(vasID, gRow, colBarcode), 1, 2) = "99" Then
            res = Insert_Data_QC(gRow)
        Else
            res = Insert_Data(gRow)
        End If

        If res = 1 Then
        
            SetText vasID, "완료", gRow, colState
            SetBackColor vasID, gRow, gRow, colBarcode, colState, 202, 255, 112
            
            vasID.Col = 1
            vasID.Row = gRow
            vasID.Value = 0
            
            SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                  " SENDFLAG = '2' " & vbCrLf & _
                  " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                  " AND BARCODE = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
            res = SendQuery(gLocal, SQL)
            'If Mid(GetText(vasID, gRow, colBarcode), 1, 2) = "99" Then
            vasID.DeleteRows gRow, 1
            ClearSpread vasRes
            vasID.MaxRows = vasID.DataRowCnt
        Else
            SetText vasID, "실패", gRow, colState
            SetForeColor vasID, gRow, gRow, colBarcode, colState, 255, 0, 0
            SetBackColor vasID, gRow, gRow, colBarcode, colState, 255, 255, 255
            
        End If
        
        gRow = -1
    
    End If
    
    Call vasID_Click(colBarcode, gRow)          '/////결과 받은후 화면에 표시
    
End Sub

Function eGFR_SAVE(asID As String, asResult As String)
    Dim i, k As Integer
    Dim lsReuslt As String
    Dim lsReuslt1 As String
    Dim sResType As String
    '//////// HL , Panic, Delta 체크
    Dim Ref_Flag        As String
    Dim Ref_Cnt         As Integer
    Dim Ref_HL          As String
    Dim Ref_Delta       As String
    Dim Ref_Panic       As String
        
    If IsNumeric(GetText(vasList, 1, colAge)) = False Then: Exit Function
    If IsNumeric(asResult) = False Then Exit Function
    If GetText(vasList, 1, colSex) = "M" Then
        lsReuslt = 186 * (CCur(asResult) ^ (-1.154)) * (CCur(GetText(vasList, 1, colAge)) ^ (-0.203))
    Else
        lsReuslt = 186 * (CCur(asResult) ^ (-1.154)) * (CCur(GetText(vasList, 1, colAge)) ^ (-0.203)) * 0.742
    End If
    
    
    'lsReuslt1 = Format(lsReuslt, "###0.0")
    
    Call EquipExamCode("eGFR", asID)
    
    If IsNumeric(gExamRange) = True Then
        For k = 0 To gExamRange
                    If k = 0 Then
                sResType = "#0"
            ElseIf k = 1 Then
                sResType = sResType & ".0"
            Else
                sResType = sResType & "0"
            End If
        Next
            
        lsReuslt1 = Format(lsReuslt, sResType)
    Else
    
    End If
     
    SQL = "Select examcode, examname, reflow, refhigh, resprec from equipexam" & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examcode  = '" & gEquipExamCode & "' " & vbCrLf & _
          "  And equipcode = 'eGFR' "
    res = db_select_Col(gLocal, SQL)                        '///// 검사 결과 저장용
    
    
    '////// HL, Delta, Panuic  체크
    Ref_Flag = GetDecision(gRow, asID, Trim(gReadBuf(0)), lsReuslt)
    Ref_Cnt = InStr(1, Ref_Flag, "/")
    Ref_HL = Mid(Ref_Flag, 1, Ref_Cnt - 1)
    
    Ref_Flag = Mid(Ref_Flag, Ref_Cnt + 1)
    Ref_Cnt = InStr(1, Ref_Flag, "/")
    Ref_Delta = Mid(Ref_Flag, 1, Ref_Cnt - 1)
    
    Ref_Flag = Mid(Ref_Flag, Ref_Cnt + 1)
    Ref_Panic = Ref_Flag
    
    
    '///// 결과 입력
    i = vasResTemp.DataRowCnt + 1
    
    SetText vasResTemp, "eGFR", i, colEquipCode     '장비코드
    SetText vasResTemp, Trim(gReadBuf(0)), i, colExamCode    '검사코드
    SetText vasResTemp, Trim(gReadBuf(1)), i, colExamName    '검사명
    SetText vasResTemp, lsReuslt1, i, colResult              '검사결과
    SetText vasResTemp, lsReuslt, i, colEquipResult            '장비 검사결과
    
    SetText vasResTemp, Ref_HL, i, colFLAG
    SetText vasResTemp, Ref_Delta, i, colDelta
    SetText vasResTemp, Ref_Panic, i, colPanic
    
    Save_Local_One_Signal 1, i, "1", vasResTemp, lsReuslt1
End Function

Function Proc_Order_LX(asBarcode As String) As Integer
    Dim i, j As Integer
    Dim sCnt As String
    Dim iCnt As Integer

    Dim lsData As String

    Dim lsFunc As String
    Dim lsSampleNo As String
    Dim lsDisk As String
    Dim lsPosNO As String
    Dim lsID As String
    Dim lsExamCode As String

    Dim lsSpcCode As String
    
    Dim retOrder As String
    Dim retHead As String
    Dim retMiddle As String
    
    Dim lsEquipCode As String
    Dim iISE As Integer

    Dim lsClass As String

    Dim eDate As String
    Dim llRow As Long

    Dim lsOrder As String
    
    Dim lsSex As String
    Dim lsAge As String
    
    Dim vTemp As String
    
    Dim iCCR As Integer
    
    Dim iTIBC As Integer
    
    Dim SpecNo As String
    
On Error GoTo ErrHandle

    lsID = asBarcode

    retOrder = ""
    lsOrder = ""
    gOrderMessage = ""
    
    eDate = Format(CDate(GetDateFull), "yyyymmdd")

    Proc_Order_LX = -1
    
    llRow = vasID.DataRowCnt + 1
    If llRow > vasID.MaxRows Then
        vasID.MaxRows = llRow + 1
    End If

    If Trim(lsID) = "" Then
        Exit Function
    End If
    SetText vasID, Trim(lsID), llRow, colBarcode
    vasActiveCell vasID, llRow, colPID

    ClearSpread vasOrder, 1, 1
    ClearSpread vasOrderBuf, 1, 1
    SetForeColor vasID, llRow, llRow, 1, colState, 0, 0, 0

    iCnt = 0

    retOrder = ""
    lsExamCode = ""
                                    
    Get_Sample_Info llRow
                                    
    SpecNo = Trim(GetText(vasID, llRow, colSpecNo))
                                    
    SQL = ""
    SQL = SQL & vbCrLf & "SELECT A.EXMN_CD "
    SQL = SQL & vbCrLf & "  FROM SPSLHRRST A "
    SQL = SQL & vbCrLf & " WHERE A.SPCM_NO = '" & Trim(SpecNo) & "' "
    SQL = SQL & vbCrLf & "   AND A.RSLT_STAT < '2' "
    SQL = SQL & vbCrLf & "   AND A.EXMN_CD IN (" & gAllExam & ") "
    SQL = SQL & vbCrLf & "GROUP BY A.EXMN_CD "
    res = db_select_Vas(gServer, SQL, vasOrderBuf)
    
    If res < 1 Then

        SetText vasID, "없음", llRow, colState

        Exit Function
    Else

        SetText vasID, lsID, llRow, colBarcode
        
    End If
    
    For i = 1 To vasOrderBuf.DataRowCnt
        If lsExamCode = "" Then
            lsExamCode = "'" & Trim(GetText(vasOrderBuf, i, 1)) & "'"
        Else
            lsExamCode = lsExamCode & ", '" & Trim(GetText(vasOrderBuf, i, 1)) & "'"
        End If
    Next i
    
    If lsExamCode <> "" Then
        ClearSpread vasTemp
        
        SQL = "SELECT EQUIPCODE, EXAMTYPE, EXAMCODE , EXAMNAME "
        SQL = SQL & vbCrLf & "  FROM EQUIPEXAM"
        SQL = SQL & vbCrLf & " WHERE EQUIPNO = '" & gEquip & "' "
        SQL = SQL & vbCrLf & "   AND EXAMCODE IN (" & lsExamCode & ") "
        SQL = SQL & vbCrLf & " GROUP BY EQUIPCODE, EXAMTYPE, EXAMCODE, EXAMNAME "
        'SQL = SQL & vbCrLf & "  "
        
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
        
        
        For i = 1 To vasTemp.DataRowCnt
            lsEquipCode = Trim(GetText(vasTemp, i, 1))
            If lsEquipCode = "TIBC" Then: lsEquipCode = "UIBX"
            If lsEquipCode <> "" Then
                If Trim(lsOrder) = "" And (lsEquipCode = "UIBX" Or IsNumeric(Mid(Trim(lsEquipCode), 1, 2)) = True) Then
                    lsOrder = SetSpace(lsEquipCode, 4, 2) & ",0"
                    If lsEquipCode = "UIBX" Then: lsEquipCode = "TIBC"
                    iCnt = iCnt + 1
                ElseIf Trim(lsOrder) <> "" And (lsEquipCode = "UIBX" Or IsNumeric(Mid(Trim(lsEquipCode), 1, 2)) = True) Then
                    lsOrder = lsOrder & "," & SetSpace(lsEquipCode, 4, 2) & ",0"
                    iCnt = iCnt + 1
                    If lsEquipCode = "UIBX" Then: lsEquipCode = "TIBC"
                End If
                
                'iCnt = iCnt + 1
                        
                If vasOrder.MaxRows < iCnt Then
                    vasOrder.MaxRows = iCnt
                End If
                
                SetText vasOrder, lsEquipCode, iCnt, colEquipCode
                SetText vasOrder, Trim(GetText(vasTemp, i, 3)), iCnt, colExamCode
                SetText vasOrder, Trim(GetText(vasTemp, i, 4)), iCnt, colExamName
                
                Save_Local_One_Order llRow, iCnt, "0"
            Else
                    
                iCnt = iCnt + 1
                
                If vasOrder.MaxRows < iCnt Then
                    vasOrder.MaxRows = iCnt
                End If
                
                SetText vasRes, Trim(GetText(vasTemp, i, 1)), iCnt, colEquipCode
                SetText vasRes, Trim(GetText(vasTemp, i, 3)), iCnt, colExamCode
                SetText vasRes, Trim(GetText(vasTemp, i, 4)), iCnt, colExamName
                
                Save_Local_One_Order llRow, iCnt, "0"
                
            End If
        Next i
    End If
    
    '=======================================================================
    'SampleType에 가져오는 부분
    SQL = "select SAMPLETYPE from equipexam where examcode in (" & lsExamCode & ") AND  SAMPLETYPE <> '' "
    res = db_select_Col(gLocal, SQL)
    
    lsClass = gReadBuf(0)
    
    '=======================================================================
    
    lsSex = Trim(GetText(vasID, llRow, colSex))
    If lsSex <> "M" And lsSex <> "F" Then
        lsSex = "M"
    End If
    lsAge = Trim(GetText(vasID, llRow, colAge))
    If Not IsNumeric(lsAge) Then
        lsAge = 5
    End If
    'lsAge = Format(lsAge, "000")
    
    retOrder = ""
    retHead = " 0,801,01,0000,00,0,RO," & lsClass & "," & SetSpace(lsID, 15, 2) & ","
    retHead = retHead & Space(20) & ","
    retHead = retHead & Space(12) & ","
    retHead = retHead & Space(25) & ","
    retHead = retHead & Space(18) & ","
    retMiddle = Space(15) & "," & Space(1) & ","
    retMiddle = retMiddle & SetSpace(lsID, 15, 2) & ","
    retMiddle = retMiddle & Space(18) & ","
    retMiddle = retMiddle & Space(8) & ","
    retMiddle = retMiddle & Space(4) & ","
    'retMiddle = retMiddle & Format(Date, "ddmmyyyy") & ","
    'retMiddle = retMiddle & Format(Time, "hhmm") & ","
    retMiddle = retMiddle & Space(20) & ","
    retMiddle = retMiddle & Space(3) & ",5," & Space(8) & ",M,"
    retMiddle = retMiddle & Space(45) & ","
    retMiddle = retMiddle & Space(7) & "," & Space(4) & "," & Space(4) & ","
    retMiddle = retMiddle & Space(2) & "," & Space(6) & ","
    retOrder = retHead & retMiddle & Format(iCnt, "000") & "," & lsOrder
    
    'retOrder = retOrder & "020,09A ,0,43B ,0,06A ,0,05A ,0,41A ,0,44A ,0,07A ,0,08A ,0,11A ,0,35A ,0,30A ,0,31A ,0,03A ,0,12A ,0,10A ,0,01A ,0,01B ,0,04A ,0,02A ,0,50A ,0"
    retOrder = "[" & retOrder & "]"
    
    Debug.Print Time & " : " & retOrder
    'retOrder = "[ 0,801,01,0000,00,0,RO,SE,1117551401     ,                    ,            ,                         ,                  ,               , ,1117551401     ,                  ,        ,    ,                    ,   , ,        , ,                                             ,       ,    ,    ,  ,      ,011,07D ,0,08D ,0,11A ,0,30A ,0,31A ,0,06D ,0,05D ,0,03E ,0,01A ,0,01B ,0,04A ,0]"
    gOrderMessage = retOrder & CS(retOrder) & Chr(13) & Chr(10)
    
    'vasTemp1.MaxRows = vasTemp1.DataRowCnt + 1
    vasOrder_Signal.AutoSize = True
    vasOrder_Signal.SetText 1, vasOrder_Signal.DataRowCnt + 1, gOrderMessage

    SetText vasID, iCnt, llRow, colOCnt
    If iCnt = 0 Then
        SetText vasID, "없음", llRow, colState
        SetForeColor vasID, llRow, llRow, 2, 2, 255, 0, 0
    Else
        SetText vasID, iCnt, llRow, colOCnt
        SetText vasID, "오더", llRow, colState
        SetForeColor vasID, llRow, llRow, 2, 2, 0, 0, 0
    End If
    SetFont vasID, llRow, llRow, 1, vasID.MaxCols, 9, False

    vasActiveCell vasID, llRow, 1

        
    Proc_Order_LX = 1
    
    Exit Function

ErrHandle:
    Proc_Order_LX = -1
    SaveQuery SQL
    Resume Next
End Function

Function Proc_Order_LX_QC(asBarcode As String) As Integer
    Dim i, j As Integer
    Dim sCnt As String
    Dim iCnt As Integer

    Dim lsData As String

    Dim lsFunc As String
    Dim lsSampleNo As String
    Dim lsDisk As String
    Dim lsPosNO As String
    Dim lsID As String
    Dim lsExamCode As String

    Dim lsSpcCode As String
    
    Dim retOrder As String
    Dim retHead As String
    Dim retMiddle As String
    
    Dim lsEquipCode As String
    Dim iISE As Integer

    Dim lsClass As String

    Dim eDate As String
    Dim llRow As Long

    Dim lsOrder As String

    Dim lsSex As String
    Dim lsAge As String
    
    Dim vTemp As String
    
    Dim iCCR As Integer
    
    Dim iTIBC As Integer
    
    Dim SpecNo As String
    
On Error GoTo ErrHandle

    lsID = asBarcode
    
    retOrder = ""
    lsOrder = ""
    gOrderMessage = ""
    
    ClearSpread vasOrder, 1, 1
    ClearSpread vasOrderBuf, 1, 1
    eDate = Format(CDate(GetDateFull), "yyyymmdd")

    Proc_Order_LX_QC = -1
    
    llRow = vasID.DataRowCnt + 1
    If llRow > vasID.MaxRows Then
        vasID.MaxRows = llRow + 1
    End If

    If Trim(lsID) = "" Then
        Exit Function
    End If
    SetText vasID, Trim(lsID), llRow, colBarcode
    vasActiveCell vasID, llRow, colPID

    ClearSpread vasOrder, 1, 1
    
    SetForeColor vasID, llRow, llRow, 1, colState, 0, 0, 0

    iCnt = 0

    retOrder = ""
    lsExamCode = ""
                                  
    Get_Sample_Info_QC llRow
    '////////// QC오더 가지고 오기
    SpecNo = Trim(GetText(vasID, llRow, colSpecNo))
    
    ClearSpread vasOrderBuf
    SQL = ""
    SQL = "SELECT QC_EXMN_CD "
    SQL = SQL & vbCrLf & " FROM SPSLMQMST "
    SQL = SQL & vbCrLf & "WHERE EQPM_CD = '" & Mid(SpecNo, 3, 3) & "' "     '//// 장비 번호
    SQL = SQL & vbCrLf & "  AND SBSN_CD = '" & Mid(SpecNo, 6, 3) & "' "     '//// 검사명 번호
    SQL = SQL & vbCrLf & "  AND LVL_CD = '" & Mid(SpecNo, 9, 1) & "' "      '//// 레벨 번호
    SQL = SQL & vbCrLf & "  AND QC_EXMN_CD IN (" & gAllExam & ") "

    res = db_select_Vas(gServer, SQL, vasOrderBuf)
    
    If res < 1 Then

        SetText vasID, "없음", llRow, colState

        Exit Function
    Else

        SetText vasID, lsID, llRow, colBarcode
        
    End If
    
    For i = 1 To vasOrderBuf.DataRowCnt
        If lsExamCode = "" Then
            lsExamCode = "'" & Trim(GetText(vasOrderBuf, i, 1)) & "'"
        Else
            lsExamCode = lsExamCode & ", '" & Trim(GetText(vasOrderBuf, i, 1)) & "'"
        End If
    Next i
    
    If lsExamCode <> "" Then
        ClearSpread vasTemp
        
        SQL = "SELECT EQUIPCODE, EXAMTYPE, EXAMCODE , EXAMNAME "
        SQL = SQL & vbCrLf & "  FROM EQUIPEXAM"
        SQL = SQL & vbCrLf & " WHERE EQUIPNO = '" & gEquip & "' "
        SQL = SQL & vbCrLf & "   AND EXAMCODE IN (" & lsExamCode & ") "
        SQL = SQL & vbCrLf & " GROUP BY EQUIPCODE, EXAMTYPE, EXAMCODE, EXAMNAME "
        'SQL = SQL & vbCrLf & "  "
        
        res = db_select_Vas(gLocal, SQL, vasTemp)
        
        
        
        For i = 1 To vasTemp.DataRowCnt
            lsEquipCode = Trim(GetText(vasTemp, i, 1))
            If lsEquipCode = "TIBC" Then: lsEquipCode = "UIBX"
            If lsEquipCode <> "" Then
                If Trim(lsOrder) = "" And (lsEquipCode = "UIBX" Or IsNumeric(Mid(Trim(lsEquipCode), 1, 2)) = True) Then
                    lsOrder = SetSpace(lsEquipCode, 4, 2) & ",0"
                    If lsEquipCode = "UIBX" Then: lsEquipCode = "TIBC"
                    iCnt = iCnt + 1
                ElseIf Trim(lsOrder) <> "" And (lsEquipCode = "UIBX" Or IsNumeric(Mid(Trim(lsEquipCode), 1, 2)) = True) Then
                    lsOrder = lsOrder & "," & SetSpace(lsEquipCode, 4, 2) & ",0"
                    iCnt = iCnt + 1
                    If lsEquipCode = "UIBX" Then: lsEquipCode = "TIBC"
                End If
                
                'iCnt = iCnt + 1
                        
                If vasOrder.MaxRows < iCnt Then
                    vasOrder.MaxRows = iCnt
                End If
                
                SetText vasOrder, lsEquipCode, iCnt, colEquipCode
                SetText vasOrder, Trim(GetText(vasTemp, i, 3)), iCnt, colExamCode
                SetText vasOrder, Trim(GetText(vasTemp, i, 4)), iCnt, colExamName
                
                Save_Local_One_Order llRow, iCnt, "0"
            Else
                    
                iCnt = iCnt + 1
                
                If vasOrder.MaxRows < iCnt Then
                    vasOrder.MaxRows = iCnt
                End If
                
                SetText vasRes, Trim(GetText(vasTemp, i, 1)), iCnt, colEquipCode
                SetText vasRes, Trim(GetText(vasTemp, i, 3)), iCnt, colExamCode
                SetText vasRes, Trim(GetText(vasTemp, i, 4)), iCnt, colExamName
                
                Save_Local_One_Order llRow, iCnt, "0"
                
            End If
        Next i
    End If
    
    '=======================================================================
    'SampleType에 가져오는 부분
    SQL = "select SAMPLETYPE from equipexam where examcode in (" & lsExamCode & ")"
    res = db_select_Col(gLocal, SQL)
    
    lsClass = gReadBuf(0)
    
    '=======================================================================
    
    lsSex = Trim(GetText(vasID, llRow, colSex))
    If lsSex <> "M" And lsSex <> "F" Then
        lsSex = "M"
    End If
    lsAge = Trim(GetText(vasID, llRow, colAge))
    If Not IsNumeric(lsAge) Then
        lsAge = 5
    End If
    'lsAge = Format(lsAge, "000")
    
    retOrder = ""
    retHead = " 0,801,01,0000,00,0,RO," & lsClass & "," & SetSpace(lsID, 15, 2) & ","
    retHead = retHead & Space(20) & ","
    retHead = retHead & Space(12) & ","
    retHead = retHead & Space(25) & ","
    retHead = retHead & Space(18) & ","
    retMiddle = Space(15) & "," & Space(1) & ","
    retMiddle = retMiddle & SetSpace(lsID, 15, 2) & ","
    retMiddle = retMiddle & Space(18) & ","
    retMiddle = retMiddle & Space(8) & ","
    retMiddle = retMiddle & Space(4) & ","
    'retMiddle = retMiddle & Format(Date, "ddmmyyyy") & ","
    'retMiddle = retMiddle & Format(Time, "hhmm") & ","
    retMiddle = retMiddle & Space(20) & ","
    retMiddle = retMiddle & Space(3) & ",5," & Space(8) & ",M,"
    retMiddle = retMiddle & Space(45) & ","
    retMiddle = retMiddle & Space(7) & "," & Space(4) & "," & Space(4) & ","
    retMiddle = retMiddle & Space(2) & "," & Space(6) & ","
    retOrder = retHead & retMiddle & Format(iCnt, "000") & "," & lsOrder
    
    'retOrder = retOrder & "020,09A ,0,43B ,0,06A ,0,05A ,0,41A ,0,44A ,0,07A ,0,08A ,0,11A ,0,35A ,0,30A ,0,31A ,0,03A ,0,12A ,0,10A ,0,01A ,0,01B ,0,04A ,0,02A ,0,50A ,0"
    retOrder = "[" & retOrder & "]"
    
    Debug.Print Time & " : " & retOrder
    'retOrder = "[ 0,801,01,0000,00,0,RO,SE,1117551401     ,                    ,            ,                         ,                  ,               , ,1117551401     ,                  ,        ,    ,                    ,   , ,        , ,                                             ,       ,    ,    ,  ,      ,011,07D ,0,08D ,0,11A ,0,30A ,0,31A ,0,06D ,0,05D ,0,03E ,0,01A ,0,01B ,0,04A ,0]"
    gOrderMessage = retOrder & CS(retOrder) & Chr(13) & Chr(10)
    
    'vasTemp1.MaxRows = vasTemp1.DataRowCnt + 1
    vasOrder_Signal.AutoSize = True
    vasOrder_Signal.SetText 1, vasOrder_Signal.DataRowCnt + 1, gOrderMessage


    SetText vasID, iCnt, llRow, colOCnt
    If iCnt = 0 Then
        SetText vasID, "없음", llRow, colState
        SetForeColor vasID, llRow, llRow, 2, 2, 255, 0, 0
    Else
        SetText vasID, iCnt, llRow, colOCnt
        SetText vasID, "오더", llRow, colState
        SetForeColor vasID, llRow, llRow, 2, 2, 0, 0, 0
    End If
    SetFont vasID, llRow, llRow, 1, vasID.MaxCols, 9, False

    vasActiveCell vasID, llRow, 1

        
    Proc_Order_LX_QC = 1
    
    Exit Function

ErrHandle:
    Proc_Order_LX_QC = -1
    SaveQuery SQL
    Resume Next
End Function

Function SetResult(asResult As String, asEquipCode As String)
    Dim i As Integer
    Dim sLVal As String
    Dim sHVal As String
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
    Dim sPoint As Integer
    Dim sResType As String
    Dim sResFlag As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    If IsNumeric(sEquipRes) = False Then
        Exit Function
    End If
    
    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
    res = db_select_Col(gLocal, SQL)
    
    If IsNumeric(gReadBuf(0)) = True Then
        sPoint = CInt(gReadBuf(0))
        sResType = ""
        For i = 0 To sPoint
            If i = 0 Then
                sResType = "#0"
            ElseIf i = 1 Then
                sResType = sResType & ".0"
            Else
                sResType = sResType & "0"
            End If
        Next
        
        sResult = Format(sEquipRes, sResType)
    Else
        sResult = sEquipRes
    End If
    
''    If IsNumeric(gReadBuf(1)) = True Then
''        sLVal = gReadBuf(1)
''        If CCur(sLVal) > CCur(sEquipRes) Then
''            sResFlag = "H"
''        End If
''    End If
''
''    If IsNumeric(gReadBuf(2)) = True Then
''        sHVal = gReadBuf(2)
''        If CCur(sHVal) < CCur(sEquipRes) Then
''            sResFlag = ">"
''        End If
''    End If
    
    If IsNumeric(gReadBuf(1)) = True And IsNumeric(gReadBuf(2)) = True Then
        sLVal = gReadBuf(1)
        sHVal = gReadBuf(2)
        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
            sResFlag = ""
        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
            sResFlag = "H"
        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
            sResFlag = "L"
        End If
    End If
    gsFlag = sResFlag
    'sResult = sResFlag & sResult
    SetResult = sResult
    
End Function

Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    Dim RCnt As Integer
    Dim OCnt As Integer
    
'    SQL = "SELECT COUNT(*) FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
'          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
'    res = db_select_Col(gLocal, SQL)

    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
'    SQL = "SELECT  MAX(RESCNT) FROM PAT_RES WHERE BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "'"
'    res = db_select_Col(gLocal, SQL)
'    If Trim(gReadBuf(0)) = "" Then
'        RCnt = 1
'    Else
'        RCnt = CCur(gReadBuf(0)) + 1
'    End If
    
    SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
          "POSNO, PID, PNAME, " & vbCrLf & _
          "PSEX, PAGE, " & vbCrLf & _
          "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
          "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, EQUIPRESULT, RECENO, SAMPLESEQ) " & vbCrLf & _
          "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, asRow1, colBarcode)) & "', '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colPos)) & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colSex)) & "', " & 0 & ", " & vbCrLf & _
          "'" & Trim(sExamDate) & "', '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', '" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
          "'" & asSend & "', '" & "''" & "', '" & Trim(asEquipResult) & "', '" & Trim(GetText(vasID, asRow1, colSpecNo)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, 0)) & "')"
    res = SendQuery(gLocal, SQL)

    
End Function

Function Save_Local_One_Signal(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, asVasRes As Object, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    Dim RCnt As Integer
    Dim OCnt As Integer

    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND BARCODE = '" & Trim(GetText(vasList, asRow1, colBarcode)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(asVasRes, asRow2, colEquipCode)) & "' " & vbCrLf & _
          "  and examcode = '" & Trim(GetText(asVasRes, asRow2, colExamCode)) & "' "
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
        
    SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
          "POSNO, PID, PNAME, " & vbCrLf & _
          "PSEX, PAGE, " & vbCrLf & _
          "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
          "SEQNO, RESULT, EXAMNAME, SENDFLAG, EQUIPRESULT, RECENO, SAMPLESEQ, RESDATE, " & vbCrLf & _
          "REFFLAG, DELTAFLAG, PANICFLAG) " & vbCrLf & _
          "VALUES('" & gEquip & "', '" & Trim(GetText(vasList, asRow1, colBarcode)) & "', '" & Trim(GetText(vasList, asRow1, colRack)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow1, colPos)) & "', '" & Trim(GetText(vasList, asRow1, colPID)) & "', '" & Trim(GetText(vasList, asRow1, colPName)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow1, colSex)) & "', " & 0 & ", " & vbCrLf & _
          "'" & Trim(sExamDate) & "', '" & Trim(GetText(asVasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(asVasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
          "'" & Trim(GetText(asVasRes, asRow2, colSeq)) & "', '" & Trim(GetText(asVasRes, asRow2, colResult)) & "', '" & Trim(GetText(asVasRes, asRow2, colExamName)) & "', " & vbCrLf & _
          "'" & asSend & "', '" & Trim(asEquipResult) & "', '" & Trim(GetText(vasList, asRow1, colSpecNo)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow1, 0)) & "', '" & Trim(GetText(vasList, asRow1, colTestType + 1)) & "', " & vbCrLf & _
          "'" & Trim(GetText(asVasRes, asRow2, colFLAG)) & "', '" & Trim(GetText(asVasRes, asRow2, colDelta)) & "', '" & Trim(GetText(asVasRes, asRow2, colPanic)) & "')"
    res = SendQuery(gLocal, SQL)
    Debug.Print SQL
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Function Save_Local_One_Signal1(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, asVasRes As Object, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    Dim RCnt As Integer
    Dim OCnt As Integer

    SQL = ""
    SQL = SQL & vbCrLf & "UPDATE PAT_RES"
    SQL = SQL & vbCrLf & "   SET DISKNO = '" & Trim(GetText(vasList, asRow1, colRack)) & "', "
    SQL = SQL & vbCrLf & "       POSNO = '" & Trim(GetText(vasList, asRow1, colPos)) & "', "
    SQL = SQL & vbCrLf & "       RESDATE = '" & Trim(GetText(vasList, asRow1, colTestType + 1)) & "' "
    SQL = SQL & vbCrLf & " WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' "
    SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
    SQL = SQL & vbCrLf & "   AND BARCODE = '" & Trim(GetText(vasList, asRow1, colBarcode)) & "' "
    
    res = SendQuery(gLocal, SQL)
    'Debug.Print SQL
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
        
    
End Function

Function Save_Local_One_Order(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    Dim RCnt As Integer
    Dim OCnt As Integer
    
'    SQL = "SELECT COUNT(*) FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  and equipcode = '" & Trim(GetText(vasOrder, asRow2, colEquipCode)) & "'" & vbCrLf & _
'          "  and examcode= '" & Trim(GetText(vasOrder, asRow2, colExamCode)) & "'"
'    res = db_select_Col(gLocal, SQL)

    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(vasOrder, asRow2, colEquipCode)) & "'" & vbCrLf & _
          "  and examcode= '" & Trim(GetText(vasOrder, asRow2, colExamCode)) & "'"
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
'    SQL = "SELECT  MAX(RESCNT) FROM PAT_RES WHERE BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "'"
'    res = db_select_Col(gLocal, SQL)
'    If Trim(gReadBuf(0)) = "" Then
'        RCnt = 1
'    Else
'        RCnt = CCur(gReadBuf(0)) + 1
'    End If
    
    SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
          "POSNO, PID, PNAME, " & vbCrLf & _
          "PSEX, PAGE, " & vbCrLf & _
          "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
          "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, EQUIPRESULT, RECENO, SAMPLESEQ) " & vbCrLf & _
          "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, asRow1, colBarcode)) & "', '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colPos)) & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colSex)) & "', " & 0 & ", " & vbCrLf & _
          "'" & Trim(sExamDate) & "', '" & Trim(GetText(vasOrder, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasOrder, asRow2, colExamCode)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasOrder, asRow2, colSeq)) & "', '" & Trim(GetText(vasOrder, asRow2, colResult)) & "', '" & Trim(GetText(vasOrder, asRow2, colExamName)) & "', " & vbCrLf & _
          "'" & asSend & "', '" & Trim(GetText(vasOrder, asRow2, 7)) & "', '" & Trim(asEquipResult) & "', '" & Trim(GetText(vasID, asRow1, colSpecNo)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, 0)) & "')"
    res = SendQuery(gLocal, SQL)
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



Private Sub picLogin_Click()

    Dim sMsg As String
    sMsg = "검사자를 입력해주세요."
    lblUser.Caption = InputBox(sMsg, "검사자 입력")

End Sub

Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    For i = BlockRow To BlockRow2
        vasID.Col = 1
        vasID.Row = i
        If vasID.Value = 0 Then
        vasID.Value = 1
        Else
        vasID.Value = 0
        End If
    Next i
End Sub

'Private Sub Picture1_Click()
'    frmUser.Show 0
'n
'End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
  
    lsID = Trim(GetText(vasID, Row, colBarcode))
    lblPnameNow.Caption = Trim(GetText(vasID, Row, colPName))
    lblBarcodeNow.Caption = Trim(GetText(vasID, Row, colBarcode))
    
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, REFFLAG, EQUIPRESULT, DELTAFLAG, PANICFLAG  " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' " & vbCrLf & _
          " AND DISKNO = '" & Trim(GetText(vasID, Row, colRack)) & "' " & vbCrLf & _
          " AND POSNO = '" & Trim(GetText(vasID, Row, colPos)) & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, REFFLAG, EQUIPRESULT, DELTAFLAG, PANICFLAG "
    
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.MaxRows = vasRes.DataRowCnt
    
    For i = 1 To vasRes.MaxRows
        If Trim(GetText(vasRes, i, colFLAG)) = "H" Then
            SetForeColor vasRes, i, i, colResult, colResult, 255, 0, 0
        ElseIf Trim(GetText(vasRes, i, colFLAG)) = "L" Then
            SetForeColor vasRes, i, i, colResult, colResult, 0, 255, 0
        End If
    Next i
    
End Sub

'Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim iRow As Long
'    Dim lsID As String
'    Dim lsTime As String
'    Dim lsPid As String
'    Dim i As Integer
'
'    iRow = vasID.ActiveRow
'    If KeyCode = vbKeyDelete Then
'        If iRow < 1 Or iRow > vasID.DataRowCnt Then
'            Exit Sub
'        End If
'
'        lsID = Trim(GetText(vasID, iRow, colBarcode))
'        lsPid = Trim(GetText(vasID, iRow, colPID))
'
'        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " AND BARCODE = '" & lsID & "' " & vbCrLf & _
'              " AND PID = '" & lsPid & "' " & vbCrLf & _
'              " AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' " & vbCrLf & _
'              " AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' " & vbCrLf & _
'              " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'        res = SendQuery(gLocal, SQL)
'
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        DeleteRow vasID, iRow, iRow
'        vasRes.MaxRows = 0
'    ElseIf KeyCode = 13 Then
'
'        Get_Sample_Info (iRow)
'
'        lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'        'Local에서 불러오기
'        ClearSpread vasTemp
'
'        '장비코드, 검사코드, 검사명, 결과, 순번
'        SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, SEQNO " & vbCrLf & _
'              "  FROM EQUIPEXAM " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " ORDER BY SEQNO "
'
'        res = db_select_Vas(gLocal, SQL, vasTemp)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'        If lsID <> lblChangeBar.Caption Then
'            For i = 1 To 3
'                SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
'                  "POSNO, PID, PNAME, " & vbCrLf & _
'                  "JUMIN, PSEX, PAGE, " & vbCrLf & _
'                  "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
'                  "SEQNO, RESULT, EXAMNAME, " & vbCrLf & _
'                  "SENDFLAG, Hospital, refflag) " & vbCrLf & _
'                  "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, iRow, colBarcode)) & "', '" & Trim(GetText(vasID, iRow, colRack)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasID, iRow, colPos)) & "', '" & Trim(GetText(vasID, iRow, colPID)) & "', '" & Trim(GetText(vasID, iRow, colPName)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasID, iRow, colJumin)) & "', '" & Trim(GetText(vasID, iRow, colSex)) & "', " & 0 & ", " & vbCrLf & _
'                  "'" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "', '" & Trim(GetText(vasID, 0, colState + (i * 2) - 1)) & "', '" & Trim(GetText(vasTemp, i, 2)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasTemp, i, 4)) & "', '" & Trim(GetText(vasID, iRow, colState + (i * 2) - 1)) & "', '" & Trim(GetText(vasTemp, i, 3)) & "', " & vbCrLf & _
'                  "'1', '" & Trim(GetText(vasID, iRow, colHospital)) & "', '" & Trim(GetText(vasID, iRow, colState + (i * 2))) & "')"
'                res = SendQuery(gLocal, SQL)
'            Next i
'
'            SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'                  " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'                  " AND BARCODE = '" & lblChangeBar.Caption & "' " & vbCrLf & _
'                  " AND PID = '" & lblChangePID.Caption & "' " & vbCrLf & _
'                  " AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' " & vbCrLf & _
'                  " AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' " & vbCrLf & _
'                  " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'            res = SendQuery(gLocal, SQL)
'
'        ElseIf lsID = lblChangeBar.Caption Then
'            For i = 1 To 3
'                SQL = "UPDATE PAT_RES "
'                SQL = SQL & vbCrLf & "   SET RESULT ='" & Trim(GetText(vasID, iRow, colState + (i * 2) - 1)) & "' "
'                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasID, iRow, colBarcode)) & "' "
'                SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
'                SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(vasTemp, i, 2)) & "' "
'                SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasID, 0, colState + (i * 2) - 1)) & "' "
'                SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(vasID, iRow, colPID)) & "' "
'                SQL = SQL & vbCrLf & "   AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' "
'                SQL = SQL & vbCrLf & "   AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' "
'                SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'                res = SendQuery(gLocal, SQL)
'            Next i
'        End If
'        SetText vasID, "Result", gRow, colState
'
'    End If
'
'
'End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsID  As String
    
    If (Row < 1) Or (Trim(GetText(vasID, Row, colBarcode)) = "" And Trim(GetText(vasID, Row, colPID)) = "") Then Exit Sub
    
    lsID = Trim(GetText(vasID, Row, colBarcode))
    lblFlagBarcode.Caption = lsID
    lblFlagPname.Caption = Trim(GetText(vasID, Row, colPName))
    
    If fraResFlag.Visible = False Then: fraResFlag.Visible = True
    
    SQL = "SELECT MESSAGE "
    SQL = SQL & vbCrLf & "  FROM PAT_RESMEMO "
    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(lsID) & "' "
    res = db_select_Vas(gLocal, SQL, vasResMemo)
    
End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
            
        vasID_Click colBarcode, lRow
    End If
End Sub

'Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
'    Dim sResDateTime As String
'    Dim sControl As String
'    Dim sLotNo As String
'
'    Dim sRefLow As String
'    Dim sRefHigh As String
'    Dim sRefFlag As String
'
'    Dim sCnt As String
'
'    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
'    'sControl = Trim(Left(asBarcode, 2))
'    'sLotNo = Trim(Mid(asBarcode, 3))
'    sControl = asBarcode
'    sRefFlag = ""
'
'    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
'          "where equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'          "  and levelname = '" & sControl & "' " & vbCrLf & _
'          "  and equipcode = '" & asExamCode & "' "
'    res = db_select_Col(gLocal, SQL)
'    If res > 0 Then
'        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
'            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
'            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
'            If CCur(sRefHigh) < CCur(asRes2) Then
'                sRefFlag = "H"
'            End If
'            If CCur(sRefLow) > CCur(asRes2) Then
'                sRefFlag = "L"
'            End If
'        End If
'    End If
'
'    sCnt = ""
'    SQL = "Select count(*) from qc_res " & vbCrLf & _
'          "where equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
'          "  and levelname = '" & sControl & "' " & vbCrLf & _
'          "  and equipcode = '" & asExamCode & "' "
'    res = db_select_Var(gLocal, SQL, sCnt)
'    If res <= 0 Then
'        SaveQuery SQL
'        db_RollBack gLocal
'        Exit Function
'    End If
'    res = db_select_Var(gLocal, SQL, sCnt)
'    If res <= 0 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'    If Not IsNumeric(sCnt) Then sCnt = "0"
'
'    If CInt(sCnt) > 0 Then
'        SQL = "delete from qc_res " & vbCrLf & _
'              "where equipno = '" & gEquip & "' " & vbCrLf & _
'              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
'              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
'              "  and levelname = '" & sControl & "' " & vbCrLf & _
'              "  and equipcode = '" & asExamCode & "' "
'        res = SendQuery(gLocal, SQL)
'        If res = -1 Then
'            'db_RollBack gLocal
'            SaveQuery SQL
'            Exit Function
'        End If
'    End If
'    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
'          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
'    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        'db_RollBack gLocal
'        SaveQuery SQL
'        Exit Function
'    End If
'
'End Function

Private Sub vasRID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Integer
    
    For i = BlockRow To BlockRow2
        vasRID.Col = 1
        vasRID.Row = i
        If vasRID.Value = 0 Then
        vasRID.Value = 1
        Else
        vasRID.Value = 0
        End If
    Next i
End Sub

Private Sub vasRID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim i As Integer
    
    If Row < 1 Or Row > vasRID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasRID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblBarcode.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasRID, Row, colPID))
    lblPname.Caption = Trim(GetText(vasRID, Row, colPName))
    lblRrow.Caption = Row
    'Local에서 불러오기
    ClearSpread vasRRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, REFFLAG, EQUIPRESULT, DELTAFLAG, PANICFLAG  " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "  AND EXAMDATE = '" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "' " & vbCrLf & _
          " AND DISKNO = '" & Trim(GetText(vasRID, Row, colRack)) & "' " & vbCrLf & _
          " AND POSNO = '" & Trim(GetText(vasRID, Row, colPos)) & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, REFFLAG, EQUIPRESULT, DELTAFLAG, PANICFLAG "
    
    res = db_select_Vas(gLocal, SQL, vasRRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasRRes.MaxRows = vasRRes.DataRowCnt
    
    For i = 1 To vasRRes.MaxRows
        If Trim(GetText(vasRRes, i, colFLAG)) = "H" Then
            SetForeColor vasRRes, i, i, colResult, colResult, 255, 0, 0
        ElseIf Trim(GetText(vasRRes, i, colFLAG)) = "L" Then
            SetForeColor vasRRes, i, i, colResult, colResult, 0, 255, 0
        End If
    Next i
End Sub

Private Sub vasRID_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsID  As String
    
    If (Row < 1) Or (Trim(GetText(vasRID, Row, colBarcode)) = "" And Trim(GetText(vasRID, Row, colPID)) = "") Then Exit Sub
    
    lsID = Trim(GetText(vasRID, Row, colBarcode))
    lblFlagBarcode.Caption = lsID
    lblFlagPname.Caption = Trim(GetText(vasRID, Row, colPName))
    
    If fraResFlag.Visible = False Then: fraResFlag.Visible = True
    
    SQL = "SELECT MESSAGE "
    SQL = SQL & vbCrLf & "  FROM PAT_RESMEMO "
    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(lsID) & "' "
    res = db_select_Vas(gLocal, SQL, vasResMemo)
End Sub

Private Sub vasRID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    Dim lsPid As String
    Dim i As Integer
    
    iRow = vasRID.ActiveRow
    
    If KeyCode = 13 Then
        
        Get_Sample_InfoR (iRow)
        
        lsID = Trim(GetText(vasRID, iRow, colBarcode))
        
        'Local에서 불러오기
        ClearSpread vasTemp
        
        '장비코드, 검사코드, 검사명, 결과, 순번
        SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf & _
              "FROM PAT_RES " & vbCrLf & _
              "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
              "  AND EXAMDATE = '" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "' " & vbCrLf & _
              "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "

        res = db_select_Vas(gLocal, SQL, vasTemp)
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        If lsID <> lblChangeBar.Caption Then
            For i = 1 To vasRRes.DataRowCnt
                SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
                  "POSNO, PID, PNAME, " & vbCrLf & _
                  " PSEX, PAGE, " & vbCrLf & _
                  "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
                  "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, RECENO, EQUIPRESULT) " & vbCrLf & _
                  "VALUES('" & gEquip & "', '" & Trim(GetText(vasRID, iRow, colBarcode)) & "', '" & Trim(GetText(vasRID, iRow, colRack)) & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasRID, iRow, colPos)) & "', '" & Trim(GetText(vasRID, iRow, colPID)) & "', '" & Trim(GetText(vasRID, iRow, colPName)) & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasRID, iRow, colSex)) & "', " & 0 & ", " & vbCrLf & _
                  "'" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "', '" & Trim(GetText(vasRRes, i, 1)) & "', '" & Trim(GetText(vasRRes, i, 2)) & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasRRes, i, 5)) & "', '" & Trim(GetText(vasRRes, i, 4)) & "', '" & Trim(GetText(vasRRes, i, 3)) & "', " & vbCrLf & _
                  "'1', '" & Trim(GetText(vasRRes, i, colFLAG)) & "','" & Trim(GetText(vasRID, iRow, colSpecNo)) & "', '" & Trim(GetText(vasRRes, i, 7)) & "')"
                res = SendQuery(gLocal, SQL)
            Next i
            
                SQL = " DELETE FROM PAT_RES " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND BARCODE = '" & lblChangeBar.Caption & "' " & vbCrLf & _
                      " AND PID = '" & lblChangePID.Caption & "' " & vbCrLf & _
                      " AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf & _
                      " AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf & _
                      " AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
                res = SendQuery(gLocal, SQL)
        ElseIf lsID = lblChangeBar.Caption Then
            For i = 1 To vasRRes.DataRowCnt
                SQL = "UPDATE PAT_RES "
                SQL = SQL & vbCrLf & "   SET RESULT ='" & Trim(GetText(vasRRes, i, 4)) & "' "
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' "
                SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
                SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(vasRRes, i, 2)) & "' "
                SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasRRes, i, 1)) & "' "
                SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(vasRID, iRow, colPID)) & "' "
                SQL = SQL & vbCrLf & "   AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' "
                SQL = SQL & vbCrLf & "   AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' "
                SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
                res = SendQuery(gLocal, SQL)
            Next i
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasRID.DataRowCnt Then
            Exit Sub
        End If
        
        lsID = Trim(GetText(vasRID, iRow, colBarcode))
        lsPid = Trim(GetText(vasRID, iRow, colPID))
            
        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
            
        SQL = " DELETE FROM PAT_RES " & vbCrLf & _
              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              " AND BARCODE = '" & lsID & "' " & vbCrLf & _
              " AND PID = '" & lsPid & "' " & vbCrLf & _
              " AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf & _
              " AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf & _
              " AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasRID, iRow, iRow
        vasRRes.MaxRows = 0
        
    End If
End Sub

Private Sub vasRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasRID.ActiveRow
        If lRow < 1 Or lRow > vasRID.DataRowCnt Then Exit Sub
            
        vasRID_Click colBarcode, lRow
    End If
End Sub

Private Sub vasRRes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then: vasRID_KeyDown KeyCode, 0
End Sub

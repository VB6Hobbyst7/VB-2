VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmInterface 
   BorderStyle     =   0  '없음
   Caption         =   " POCT Interface "
   ClientHeight    =   10680
   ClientLeft      =   960
   ClientTop       =   660
   ClientWidth     =   15165
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
   ScaleHeight     =   534
   ScaleMode       =   0  '사용자
   ScaleWidth      =   758.25
   Begin VB.Frame fraResFlag 
      Height          =   3480
      Left            =   15240
      TabIndex        =   50
      Top             =   6120
      Visible         =   0   'False
      Width           =   8385
      Begin FPSpread.vaSpread vasResMemo 
         Height          =   2850
         Left            =   45
         TabIndex        =   51
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
         TabIndex        =   55
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
         TabIndex        =   54
         Top             =   225
         Width           =   945
      End
      Begin VB.Label lblFlagBarcode 
         Caption         =   "1234567890ab"
         Height          =   165
         Left            =   1620
         TabIndex        =   53
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
         Left            =   135
         TabIndex        =   52
         Top             =   225
         Width           =   1380
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   2655
      Left            =   15240
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   8205
      Begin FPSpread.vaSpread vasPatList 
         Height          =   150
         Left            =   5175
         TabIndex        =   58
         Top             =   2070
         Width           =   960
         _Version        =   393216
         _ExtentX        =   1693
         _ExtentY        =   265
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
      Begin FPSpread.vaSpread vasOrderBuf 
         Height          =   600
         Left            =   855
         TabIndex        =   57
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
         SpreadDesigner  =   "frmInterface.frx":1F8B
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   510
         Left            =   1935
         TabIndex        =   56
         Top             =   1395
         Width           =   1995
         _Version        =   393216
         _ExtentX        =   3519
         _ExtentY        =   900
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
         SpreadDesigner  =   "frmInterface.frx":21A3
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   7200
         TabIndex        =   48
         Top             =   315
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
         SpreadDesigner  =   "frmInterface.frx":23BB
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   46
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
         TabIndex        =   35
         Top             =   1380
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   120
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   735
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1620
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   1980
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   960
         Left            =   6705
         TabIndex        =   30
         Top             =   1350
         Visible         =   0   'False
         Width           =   1335
         Begin MSCommLib.MSComm MSComm1 
            Left            =   135
            Top             =   270
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
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   975
         Left            =   6780
         TabIndex        =   29
         Top             =   240
         Width           =   315
         _Version        =   393216
         _ExtentX        =   556
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
         SpreadDesigner  =   "frmInterface.frx":25D3
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   3195
         TabIndex        =   36
         Top             =   180
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
         SpreadDesigner  =   "frmInterface.frx":27EB
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   4980
         TabIndex        =   37
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
         SpreadDesigner  =   "frmInterface.frx":2A03
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   1800
         TabIndex        =   38
         Top             =   180
         Width           =   1365
         _Version        =   393216
         _ExtentX        =   2408
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
         SpreadDesigner  =   "frmInterface.frx":2C1B
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   4860
         TabIndex        =   40
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   5700
         TabIndex        =   39
         Top             =   1410
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   2085
      Left            =   15210
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   9465
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   1260
         TabIndex        =   26
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
         SpreadDesigner  =   "frmInterface.frx":2E33
      End
      Begin FPSpread.vaSpread vasPrintBuf 
         Height          =   1245
         Left            =   120
         TabIndex        =   27
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
         SpreadDesigner  =   "frmInterface.frx":48AC
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   9315
      Left            =   75
      TabIndex        =   6
      Top             =   840
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   16431
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
      TabPicture(0)   =   "frmInterface.frx":4AC4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":4AE0
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
            TabIndex        =   41
            Top             =   630
            Width           =   6015
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   47
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Left            =   4200
               TabIndex        =   45
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
               TabIndex        =   44
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Left            =   1605
               TabIndex        =   43
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
               TabIndex        =   42
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "EXCEL"
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
            Left            =   13050
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   1395
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
            Top             =   240
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
            Format          =   63111168
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
            SpreadDesigner  =   "frmInterface.frx":4AFC
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   7305
            Left            =   8460
            TabIndex        =   49
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
            SpreadDesigner  =   "frmInterface.frx":5548
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8775
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   14625
         Begin VB.CommandButton Command1 
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
            Left            =   120
            TabIndex        =   59
            Top             =   300
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Command16"
            Height          =   435
            Left            =   6075
            TabIndex        =   12
            Top             =   4950
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTest 
            Height          =   675
            Left            =   1800
            TabIndex        =   11
            Top             =   4950
            Visible         =   0   'False
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
            Left            =   13050
            TabIndex        =   16
            Top             =   270
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
            Left            =   11550
            TabIndex        =   15
            Top             =   270
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
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   8235
            _Version        =   393216
            _ExtentX        =   14526
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
            SpreadDesigner  =   "frmInterface.frx":935F
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   7800
            Left            =   8490
            TabIndex        =   13
            Top             =   735
            Width           =   5985
            _Version        =   393216
            _ExtentX        =   10557
            _ExtentY        =   13758
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
            Protect         =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":9E05
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
      Top             =   10305
      Width           =   15165
      _ExtentX        =   26749
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
            TextSave        =   "2012-12-12"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오전 9:32"
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
      Width           =   15045
      _Version        =   65536
      _ExtentX        =   26538
      _ExtentY        =   1138
      _StockProps     =   15
      Caption         =   "     POCT  INTERFACE"
      ForeColor       =   16777215
      BackColor       =   11494691
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   14.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         Picture         =   "frmInterface.frx":DC1B
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
         Format          =   63111168
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
Const colEquipName = 4 '///// POCT 장비 검사위치
Const colPos = 5
Const colPID = 6
Const colPName = 7
Const colSex = 8
Const colAge = 9
Const colOCnt = 10
Const colRCnt = 11
Const colState = 12
Const colA1c = 13
Const colIFCC = 15
Const coleAg = 17

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

Dim gRow As Long
Dim gsBarCode As String
Dim gsSampleType As String
Dim gsPID As String
Dim gsRackNo As String
Dim gsEquipName As String
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
            
            If IsNumeric(GetText(vasID, lRow, colBarcode)) = False Then
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
          "  AND SENDFLAG IN ('1', '2') " & vbCrLf & _
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
'        Case "0"
'            SetText vasID, "오더", iRow, colState
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
            If IsNumeric(GetText(vasRID, lRow, colBarcode)) = False Then
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
                      " WHERE EQUIPNO = '" & gEquipCode & "' " & vbCrLf & _
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

Private Sub lblclear_Click()
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
        Case chrSTX
            CharData = lsChar
            
        Case chrETX
            CharData = CharData & lsChar
            'SaveData "[RX]" & txtdata
    
        Case chrLF
            CharData = CharData & lsChar
            SaveData "[Rx]" & CharData
            
            POCT_ILSAN CharData
            
            SaveData "[Tx]" & chrACK
            CharData = ""
            
        Case chrENQ
            
            If dtpToday <> Format(Date, "yyyy/mm/dd") Then
                dtpToday = Format(Date, "yyyy/mm/dd")
            End If
            
            
            SaveData "[RX]" & chrENQ
            
            SaveData "[TX]" & chrACK
            
        Case chrACK
            SaveData "[RX]" & chrACK
            
            gOrdRow = gOrdRow + 1
    
            If GetText(vasOrder, gOrdRow, 1) = "" Then
                Exit Sub
            End If
    
            If gOrdRow <= vasOrder.DataRowCnt Then
    
                sSendData = Trim(GetText(vasOrder, gOrdRow, 1))
    
                SaveData "[Tx]" & sSendData
    
                If gOrdRow = vasOrder.DataRowCnt Then
                    'ClearSpread vasOrderBuff
                    ClearSpread vasOrder
    
                    Me.MousePointer = 0
                End If
            End If
        Case chrEOT     '자료수신 완료
            If gRecodeType = "R" Then
                gSndState = "R"
            ElseIf gRecodeType = "Q" Then
                gOrdRow = 0
                gPreMsg = chrENQ
            
                SaveData "[Tx]" & chrENQ
                        
                gSndState = "Q"
                gPreMsg = chrENQ
            End If
            
            gMsgFlag = ""
            gHeadRecode = ""
            txtData.Text = ""
            
        Case Else
            CharData = CharData & lsChar
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

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape And fraResFlag.Visible = True Then
        fraResFlag.Visible = False
    End If
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
    Dim lsETB As Boolean
    Dim lsETBFirst As Boolean
    Dim sSendData As String
    
    dtpToday = Date
    lsChar = MSComm1.Input
    
    Select Case lsChar
    Case chrSTX
        txtData.Text = lsChar
        
    Case chrETX
        txtData.Text = txtData.Text & lsChar
        'SaveData "[RX]" & txtdata
        If lsETB = True Then: lsETB = False
            
    Case chrETB
        lsETB = True
        
    Case chrLF
        If lsETB = False Then
            txtData.Text = txtData.Text & lsChar
            SaveData "[Rx]" & txtData.Text
            
            POCT_ILSAN txtData.Text
            
            MSComm1.Output = chrACK
            SaveData "[Tx]" & chrACK
        End If
    Case chrENQ
        
        If dtpToday <> Format(Date, "yyyy/mm/dd") Then
            dtpToday = Format(Date, "yyyy/mm/dd")
        End If
        
        
        SaveData "[RX]" & chrENQ
        
        MSComm1.Output = chrACK
        SaveData "[TX]" & chrACK
        
    Case chrACK
        SaveData "[RX]" & chrACK
        
        gOrdRow = gOrdRow + 1

        If GetText(vasOrder, gOrdRow, 1) = "" Then
            Exit Sub
        End If

        If gOrdRow <= vasOrder.DataRowCnt Then

            sSendData = Trim(GetText(vasOrder, gOrdRow, 1))

            MSComm1.Output = sSendData
            SaveData "[Tx]" & sSendData

            If gOrdRow = vasOrder.DataRowCnt Then
                ClearSpread vasOrderBuf
                ClearSpread vasOrder

                Me.MousePointer = 0
            End If
        End If
    Case chrEOT     '자료수신 완료
        If gRecodeType = "R" Then
            gSndState = "R"
        ElseIf gRecodeType = "Q" Then
            gOrdRow = 0
            gPreMsg = chrENQ
            
            frmInterface.MSComm1.Output = chrENQ
            SaveData "[Tx]" & chrENQ
                    
            gSndState = "Q"
            gPreMsg = chrENQ
        End If
        
        gMsgFlag = ""
        gHeadRecode = ""
        txtData.Text = ""
        
    Case Else
    
        If lsETB = False Then
            txtData.Text = txtData.Text & lsChar
        End If
    End Select

End Sub

Sub POCT_ILSAN(asData As String)
'ASTM

    Dim MyVar As String
    Dim MyRet As String
    
    Dim i, k As Integer
    Dim j As Integer
    Dim iCnt As Integer
    Dim jCnt As Integer
    Dim aCnt As Integer
    Dim bCnt As Integer
    Dim ii  As Integer
    
    Dim iRow As Integer
    Dim lRow As Integer
    Dim liRet As Integer
    
    Dim lsDistinctII As String
    Dim lsInqueryMode As String
    Dim lsDate      As String
    Dim lsTime      As String
    Dim lsRack      As String
    Dim lsTube      As String
    Dim lsID        As String
    Dim lsIDInfo    As String
    Dim lsPName     As String
    
    Dim lsTemp      As String
    Dim lsData      As String
    
    Dim lsEquip     As String
    Dim lsTestID    As String
    Dim lsResult    As String
    Dim lsResult1   As String
    Dim lsFlag      As String
    
    Dim lsTemp1     As String
    Dim lsMessage   As String
    
    Dim lsExamCode  As String
    Dim lsRsCode    As String
    Dim lsExamName  As String
    Dim lsPoint     As String
    Dim sTmpStr     As String
    Dim lsSelExam   As String
    
    Dim sDate       As String
    Dim iExamCnt    As Integer
    
    Dim sLen, sLen2 As String
    
    Dim Data_Array              '/데이터 배열 만들기
    
    '//////// HL , Panic, Delta 체크
    Dim Ref_Flag        As String
    Dim Ref_Cnt         As Integer
    Dim Ref_HL          As String
    Dim Ref_Delta       As String
    Dim Ref_Panic       As String
    
    sDate = Format(dtpToday, "yyyymmdd")

    j = 1
    
    Select Case Mid(asData, 3, 1)
    Case "H"    'Header
        
        lsMessage = ""
        gsEquipName = ""
        ClearSpread vasRes
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        lsEquip = Mid(asData, 1, sLen - 1)
        
        sLen = InStr(1, lsEquip, "^")
        gsEquipName = Mid(lsEquip, sLen + 1)

    Case "P"    'Patient
        Data_Array = Split(asData, "|")
        
        If UBound(Data_Array) > 2 Then
            gsBarCode = Data_Array(3)
            sLen = InStr(1, gsBarCode, chrCR)
            If sLen > 2 Then
                gsBarCode = Mid(gsBarCode, 1, sLen - 1)
            Else
                gsBarCode = ""
            End If
        Else
            gsBarCode = ""
        End If
        
'        sLen = InStr(1, asData, "|")
'        asData = Mid(asData, sLen + 1)
'
'        sLen = InStr(1, asData, "|")
'        asData = Mid(asData, sLen + 1)
'
'        sLen = InStr(1, asData, "|")
'        asData = Mid(asData, sLen + 1)
'
'        sLen = InStr(1, asData, chrCR)
'        gsBarCode = Mid(asData, 1, sLen - 1)
        
'        sLen = InStr(1, gsBarCode, "^")
'        gsRackNo = Trim(Mid(gsBarCode, 1, sLen - 1))      'Rack
'        gsBarCode = Mid(gsBarCode, sLen + 1)
'
'        sLen = InStr(1, gsBarCode, "^")
'        gsPosNo = Trim(Mid(gsBarCode, 1, sLen - 1))        'Tube
'        If Len(gsPosNo) = 1 Then
'            gsPosNo = Format(gsPosNo, "0#")
'        End If
'
'        gsBarCode = Mid(gsBarCode, sLen + 1)
'
'        sLen = InStr(1, gsBarCode, "^")
        
'        If sLen <> 0 Then
'            gsBarCode = Trim(Mid(gsBarCode, 1, sLen - 1))    '검체번호
'        Else
'
'            gsBarCode = Trim(gsBarCode)    '검체번호
'        End If

        gsBarCode = Trim(gsBarCode)    '검체번호

    Case "Q"    'Request
        gRecodeType = "Q"
        
        ClearSpread vasTemp
        ClearSpread vasOrder
        'ClearSpread vasOrderBuf
        
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        asData = Mid(asData, sLen + 1)
        
        sLen = InStr(1, asData, "|")
        gsBarCode = Mid(asData, 1, sLen - 1)
        
        
        sLen = InStr(1, gsBarCode, "^")
        gsRackNo = Trim(Mid(gsBarCode, 1, sLen - 1))      'Rack
        gsBarCode = Mid(gsBarCode, sLen + 1)
        
        sLen = InStr(1, gsBarCode, "^")
        gsPosNo = Trim(Mid(gsBarCode, 1, sLen - 1))        'Tube
        If Len(gsPosNo) = 1 Then
            gsPosNo = Format(gsPosNo, "0#")
        End If
        
        gsBarCode = Mid(gsBarCode, sLen + 1)
        
        sLen = InStr(1, gsBarCode, "^")
        gsBarCode = Trim(Mid(gsBarCode, 1, sLen - 1))    '검체번호
        'gAttribute = Mid(gsBarCode, sLen + 1)           'Attribute
    
        glRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < glRow + 1 Then
            vasID.MaxRows = glRow + 1
        End If
        
        glRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                glRow = i
                Exit For
            End If
        Next i
        
        '2004/06/16 이상은========================================================
        'Order 전송뒤 Clear시 다시 바코드 스캔 안하고 결과 넘어오도록 수정
        If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
            glRow = vasID.DataRowCnt + 1
            If glRow > vasID.MaxRows Then
                vasID.MaxRows = glRow + 1
            End If
            vasActiveCell vasID, glRow, colBarcode
            SetText vasID, gsEquipName, glRow, colEquipName
            SetText vasID, gsBarCode, glRow, colBarcode
            SetText vasID, gsRackNo, glRow, colRack
            SetText vasID, gsPosNo, glRow, colPos
        End If
        '==========================================================================
                   
        '환자정보 가져오기
'        If Trim(GetText(vasID, glRow, colPID)) = "" Then
'            Get_Sample_Info glRow
'        End If
'
        If Trim(GetText(vasID, glRow, colPID)) = "" And Mid(gsBarCode, 1, 2) <> "QC" Then
            Get_Sample_Info glRow
        Else
            Get_Sample_Info_QC glRow
        End If
        
        'Order 만들기
        'Make_Order_ASTM gsBarCode, glRow
                     
    Case "O"    'Test Order
        
        
        Data_Array = Split(asData, "|")
        If gsBarCode = "" Then
            
            gsBarCode = Data_Array(15)
            
            If gsBarCode <> "" And gsBarCode <> "^" Then
                sLen = InStr(1, gsBarCode, "^")
                gsBarCode = gsEquipName & "/" & Mid(gsBarCode, 1, sLen - 1)
            Else
                gsBarCode = Data_Array(3)
            End If
        End If
        
        glRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                glRow = i
                Exit For
            End If
        Next i
        
        '2004/06/16 이상은========================================================
        'Order 전송뒤 Clear시 다시 바코드 스캔 안하고 결과 넘어오도록 수정
        If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
            glRow = vasID.DataRowCnt + 1
            If glRow > vasID.MaxRows Then
                vasID.MaxRows = glRow + 1
            End If
        End If
        '==========================================================================
        
        vasActiveCell vasID, glRow, colBarcode
        SetText vasID, gsEquipName, glRow, colEquipName
        SetText vasID, gsBarCode, glRow, colBarcode
        
        
        If Trim(gsBarCode) <> "" Then
            '환자정보 가져오기
            If Trim(GetText(vasID, glRow, colPID)) = "" And IsNumeric(gsBarCode) = True Then
                Get_Sample_Info glRow
                POCT_XML (GetText(vasID, glRow, colBarcode))
            Else
                Get_Sample_Info_QC glRow
            End If
        Else
            SetText vasID, "NOBARCODE", glRow, colBarcode
        End If
        
    Case "R"    'Result
        gRecodeType = "R"

        SetText vasID, "Result", glRow, colState
        
        If vasRes.MaxRows = 0 Then
            vasRes.MaxRows = 1
            iRow = vasRes.MaxRows
        Else
            vasRes.MaxRows = vasRes.MaxRows + 1
            iRow = vasRes.MaxRows
        End If
        
        iCnt = 0
        i = InStr(1, asData, "|")
        Do While i > 0
            iCnt = iCnt + 1
            
            lsTemp = Mid(asData, 1, i - 1)
            asData = Mid(asData, i + 1)
            
            Select Case iCnt
            Case 2  '결과갯수
                'gRCnt = Trim(lsTemp)
                
            Case 3  'TestID
                If lsTemp <> "" Then
                    lsTemp = Mid(lsTemp, 4)
                    j = InStr(1, lsTemp, "^")
                    
                    If j > 0 Then
                        lsTestID = Left(lsTemp, j - 1)
                        gsTestID = lsTestID
                    Else
                        lsTestID = Trim(lsTemp)
                        gsTestID = lsTestID
                    End If
                    
                    If lsTestID = "" Then
                        Exit Sub
                    End If
                End If
    
            Case 4 '결과 Data
                
                If IsNumeric(gsBarCode) = True Then
                    Call EquipExamCode(gsTestID, gsBarCode)
                Else
                    Call EquipExamCode_QC(gsTestID, Trim(GetText(vasID, glRow, colSpecNo)))
                End If
                gReadBuf(0) = ""
                SQL = "Select ExamCode, ExamName, resprec From EquipExam" & vbCrLf & _
                      " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                      "  And  EquipCode = '" & Trim(gsTestID) & "'" & vbCrLf & _
                      "  And  ExamCode = '" & gEquipExamCode & "' "
                res = db_select_Col(gLocal, SQL)
                
                
                If res = 1 And gReadBuf(0) <> "" Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                    'lsPoint = Trim(gReadBuf(2))
                    
                    j = vasRes.DataRowCnt + 1
                    
                    lsResult = Trim(lsTemp)
                    
                    If IsNumeric(gExamRange) = True Then
                        For k = 0 To gExamRange
                            If k = 0 Then
                                lsPoint = "#0"
                            ElseIf k = 1 Then
                                lsPoint = lsPoint & ".0"
                            Else
                                lsPoint = lsPoint & "0"
                            End If
                        Next k
                    
                        lsResult1 = Format(lsResult, lsPoint)
                    Else
                        lsResult1 = lsResult
                    End If
                
                    Ref_Flag = GetDecision(glRow, gsBarCode, Trim(gReadBuf(0)), lsResult)
                    Ref_Cnt = InStr(1, Ref_Flag, "/")
                    Ref_HL = Mid(Ref_Flag, 1, Ref_Cnt - 1)
                    
                    Ref_Flag = Mid(Ref_Flag, Ref_Cnt + 1)
                    Ref_Cnt = InStr(1, Ref_Flag, "/")
                    Ref_Delta = Mid(Ref_Flag, 1, Ref_Cnt - 1)
                    
                    Ref_Flag = Mid(Ref_Flag, Ref_Cnt + 1)
                    Ref_Panic = Ref_Flag
                
'                    If gsTestID = "WBC" Then
'                        lsResult = Format(lsResult, "#0.0")
'                    End If
                    
                    If IsNumeric(lsResult) Then
                    
                        SetText vasRes, gsTestID, j, colEquipCode               '장비코드
                        SetText vasRes, lsExamCode, j, colExamCode              '검사코드
                        SetText vasRes, lsExamName, j, colExamName              '검사명
                        SetText vasRes, lsResult1, j, colResult                  '검사결과
                        
                        SetText vasRes, Ref_HL, j, colFLAG
                        SetText vasRes, Ref_Delta, j, colDelta
                        SetText vasRes, Ref_Panic, j, colPanic
                        
                        Save_Local_One glRow, j, "1", lsResult
                    Else
                        '2004/06/09 이상은
                        'SetText vasRes, "", j, colResult
                        '================================================================
                        '결과값 없어도 항목 디스플레이 되도록
                        SetText vasRes, gsBarCode, j, colBarcode                '검체번호
                        SetText vasRes, gsTestID, j, colEquipCode               '장비코드
                        SetText vasRes, lsExamCode, j, colExamCode              '검사코드
                        SetText vasRes, lsExamName, j, colExamName              '검사명
                        SetText vasRes, lsResult1, j, colResult                        '검사결과
                            
                        Save_Local_One glRow, j, "1", lsResult
                        '================================================================
                    End If
                Else
                    gReadBuf(0) = ""
                    SQL = "Select ExamCode, ExamName From EquipExam" & vbCrLf & _
                          " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                          "  And  EquipCode = '" & Trim(gsTestID) & "'"
                    res = db_select_Col(gLocal, SQL)
                    If res = 1 Then
                        j = vasRes.DataRowCnt + 1
                        
                        lsResult = Trim(lsTemp)
                    
                        SetText vasRes, gsBarCode, j, colBarcode                '검체번호
                        SetText vasRes, gsTestID, j, colEquipCode               '장비코드
                        SetText vasRes, "", j, colExamCode              '검사코드
                        SetText vasRes, Trim(gReadBuf(1)), j, colExamName              '검사명
                        SetText vasRes, lsResult, j, colResult                  '검사결과
                        
                        Save_Local_One glRow, j, "1", lsResult
                    End If
                End If
            
            End Select
     
            lsTemp = ""
            lsTestID = ""
            lsResult = ""
            i = InStr(1, asData, "|")
        Loop
        
    Case "L"
        If glRow <> -1 And gRecodeType = "R" Then
            If MnTransAuto.Checked = True Then
                vasID.Col = 1
                vasID.Row = glRow
                vasID.Value = 1
                If IsNumeric(GetText(vasID, glRow, colBarcode)) = False Then
                    res = Insert_Data_QC(glRow)
                Else
                    res = Insert_Data(glRow)
                End If
                
                If res = 1 Then
                    SetBackColor vasID, glRow, glRow, colCheckBox, colState, 202, 255, 112
                    SetText vasID, "완료", glRow, colState
                        SQL = "update pat_res set sendflag = '2' where examdate = '" & Format(dtpToday, "YYYYMMDD") & "' " & CR & _
                              "and equipno = '" & gEquip & "' And barcode = '" & Trim(GetText(vasID, glRow, colBarcode)) & "' "
                        res = SendQuery(gLocal, SQL)
                        
                        vasID.Row = glRow
                        vasID.Col = 1
                        
                        vasID.Value = 0
                ElseIf res = -1 Then
                    SetForeColor vasID, glRow, glRow, colCheckBox, colState, 255, 0, 0
                    SetText vasID, "실패", glRow, colState
                End If
            End If
        End If
    End Select
    
End Sub


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
          "SEQNO, RESULT, EXAMNAME, SENDFLAG, EQUIPRESULT, RECENO, SAMPLESEQ, RESDATE, " & vbCrLf & _
          "REFFLAG, DELTAFLAG, PANICFLAG ) " & vbCrLf & _
          "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, asRow1, colBarcode)) & "', '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colPos)) & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colSex)) & "', " & 0 & ", " & vbCrLf & _
          "'" & Trim(sExamDate) & "', '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', '" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
          "'" & asSend & "', '" & Trim(asEquipResult) & "', '" & Trim(GetText(vasID, asRow1, colSpecNo)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, 0)) & "', '', " & vbCrLf & _
          "'" & Trim(GetText(vasRes, asRow2, colFLAG)) & "', '" & Trim(GetText(vasRes, asRow2, colDelta)) & "', '" & Trim(GetText(vasRes, asRow2, colPanic)) & "')"
          
          
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
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    'Local에서 불러오기
    ClearSpread vasRes

    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, REFFLAG, DELTAFLAG, PANICFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' " & vbCrLf & _
          " AND DISKNO = '" & Trim(GetText(vasID, Row, colRack)) & "' " & vbCrLf & _
          " AND POSNO = '" & Trim(GetText(vasID, Row, colPos)) & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, REFFLAG, DELTAFLAG, PANICFLAG  " & vbCrLf & _
          "ORDER BY EXAMCODE, EQUIPCODE "
        
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
        SetForeColor vasRes, i, i, colResult, colResult, 0, 0, 255
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

'Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim lsID  As String
'
'    If (Row < 1) Or (Trim(GetText(vasID, Row, colBarcode)) = "" And Trim(GetText(vasID, Row, colPID)) = "") Then Exit Sub
'
'    lsID = Trim(GetText(vasID, Row, colBarcode))
'    lblFlagBarcode.Caption = lsID
'    lblFlagPname.Caption = Trim(GetText(vasID, Row, colPName))
'
'    If fraResFlag.Visible = False Then: fraResFlag.Visible = True
'
'    SQL = "SELECT MESSAGE "
'    SQL = SQL & vbCrLf & "  FROM PAT_RESMEMO "
'    SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(lsID) & "' "
'    res = db_select_Vas(gLocal, SQL, vasResMemo)
'
'End Sub

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
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, REFFLAG, EQUIPRESULT, PANICFLAG, DELTAFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          " AND DISKNO = '" & Trim(GetText(vasRID, Row, colRack)) & "' " & vbCrLf & _
          " AND POSNO = '" & Trim(GetText(vasRID, Row, colPos)) & "' " & vbCrLf & _
          " AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, REFFLAG , EQUIPRESULT, PANICFLAG, DELTAFLAG " & vbCrLf & _
          "ORDER BY EXAMCODE, EQUIPCODE "
    
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
            SetForeColor vasRRes, i, i, colResult, colResult, 0, 0, 255
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
        
        If MsgBox("Barcode 번호를 변경하시겠습니까? " & vbCrLf & vbCrLf & _
              lblChangeBar.Caption & " -> " & lsID, vbQuestion + vbOKCancel + vbDefaultButton2) = vbCancel Then Exit Sub
        Call POCT_XML(lsID)
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
                If Mid(Trim(GetText(vasRID, iRow, colBarcode)), 1, 2) <> "99" And IsNumeric(Trim(GetText(vasRID, iRow, colBarcode))) = True Then
                    Call EquipExamCode(Trim(GetText(vasRRes, i, 1)), GetText(vasRID, iRow, colBarcode))
                Else
                    Call EquipExamCode_QC(Trim(GetText(vasRRes, i, 1)), GetText(vasRID, iRow, colBarcode))
                End If
                
                SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
                  "POSNO, PID, PNAME, " & vbCrLf & _
                  " PSEX, PAGE, " & vbCrLf & _
                  "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
                  "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, RECENO, EQUIPRESULT) " & vbCrLf & _
                  "VALUES('" & gEquip & "', '" & Trim(GetText(vasRID, iRow, colBarcode)) & "', '" & Trim(GetText(vasRID, iRow, colRack)) & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasRID, iRow, colPos)) & "', '" & Trim(GetText(vasRID, iRow, colPID)) & "', '" & Trim(GetText(vasRID, iRow, colPName)) & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasRID, iRow, colSex)) & "', " & 0 & ", " & vbCrLf & _
                  "'" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "', '" & Trim(GetText(vasRRes, i, 1)) & "', '" & Trim(gEquipExamCode) & "', " & vbCrLf & _
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
                If Mid(Trim(GetText(vasRID, iRow, colBarcode)), 1, 2) <> "99" And IsNumeric(Trim(GetText(vasRID, iRow, colBarcode))) = True Then
                    Call EquipExamCode(Trim(GetText(vasRRes, i, 1)), GetText(vasRID, iRow, colBarcode))
                Else
                    Call EquipExamCode_QC(Trim(GetText(vasRRes, i, 1)), GetText(vasRID, iRow, colBarcode))
                End If
                
                SQL = "UPDATE PAT_RES "
                SQL = SQL & vbCrLf & "   SET RESULT ='" & Trim(GetText(vasRRes, i, 4)) & "' "
                SQL = SQL & vbCrLf & "      ,EXAMCODE = '" & gEquipExamCode & "' "
                SQL = SQL & vbCrLf & " WHERE BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' "
                SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
                SQL = SQL & vbCrLf & "   AND EXAMCODE = '" & Trim(GetText(vasRRes, i, 2)) & "' "
                SQL = SQL & vbCrLf & "   AND EQUIPCODE = '" & Trim(GetText(vasRRes, i, 1)) & "' "
                SQL = SQL & vbCrLf & "   AND PID = '" & Trim(GetText(vasRID, iRow, colPID)) & "' "
                SQL = SQL & vbCrLf & "   AND DISKNO = '" & Trim(GetText(vasRID, iRow, colRack)) & "' "
                'SQL = SQL & vbCrLf & "   AND POSNO = '" & Trim(GetText(vasRID, iRow, colPos)) & "' "
                SQL = SQL & vbCrLf & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
                res = SendQuery(gLocal, SQL)
                Debug.Print SQL
            Next i
        End If
        Call vasRID_Click(colBarcode, iRow)
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


Function POCT_XML(asBarcode As String) As Boolean
        
  Dim XMLRequest            As New XMLHTTPRequest
  Dim txtSendXML            As String
  Dim txtResponseHeaders    As String
  Dim txtResponse           As String
  
  Dim BLPS_ID               As String '/채혈자 아이디
  
  POCT_XML = False
    SQL = "SELECT COUNT(EXMN_CD), BLPS_ID FROM SPSLMJBDI "
    SQL = SQL & vbCrLf & " WHERE SPCM_NO = (SELECT FN_LABCVTBCNO('" & Trim(asBarcode) & "') FROM DUAL)"                                             '검체번호"
    SQL = SQL & vbCrLf & "   AND EXMN_CD IN (" & gAllExam & ") "                                              '검사코드"
    SQL = SQL & vbCrLf & "   AND RSLT_STAT < '2' "                                                          '결과상태"
    SQL = SQL & vbCrLf & " GROUP BY BLPS_ID "
    res = db_select_Col(gServer, SQL)
    If gReadBuf(0) = "" Then gReadBuf(0) = "0"
    If gReadBuf(0) = "0" Then Exit Function
    BLPS_ID = gReadBuf(1)
    
On Error GoTo err_handler
    
    txtSendXML = "<?xml version='1.0' encoding='UTF-8'?>"
    txtSendXML = txtSendXML & vbCrLf & "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>"
    txtSendXML = txtSendXML & vbCrLf & "<soapenv:Body>"
    txtSendXML = txtSendXML & vbCrLf & "<registSpcmRcpn xmlns='http://svc.poct.ws.nhimc/'>"
    txtSendXML = txtSendXML & vbCrLf & "<arg0 xmlns=''>" & Trim(asBarcode) & "</arg0>"
    txtSendXML = txtSendXML & vbCrLf & "<arg1 xmlns=''>" & BLPS_ID & "</arg1>"
    'txtSendXML = txtSendXML & vbCrLf & "<arg1 xmlns=''>POCT</arg1>"
    txtSendXML = txtSendXML & vbCrLf & "</registSpcmRcpn>"
    txtSendXML = txtSendXML & vbCrLf & "</soapenv:Body>"
    txtSendXML = txtSendXML & vbCrLf & "</soapenv:Envelope>" & vbCrLf
    '/운영
    'XMLRequest.Open "POST", "http://192.168.1.101:8800/service/PoctService?wsdl", False
    XMLRequest.Open "POST", "http://isis.nhimc:8800/service/PoctService?wsdl", False
    '/테스트
    'XMLRequest.Open "POST", "http://192.168.1.20:8800/service/PoctService?wsdl", False
    
    XMLRequest.setRequestHeader "Content-Type", "text/xml"
    '  o.setRequestHeader "Connection", "close"
    XMLRequest.setRequestHeader "Connection", "PoctService"
    XMLRequest.setRequestHeader "SOAPAction", ""
    XMLRequest.send txtSendXML
        
    
    txtResponseHeaders = XMLRequest.getAllResponseHeaders
    txtResponse = XMLRequest.responseText
    Save_Raw_Data "[XML]  " & txtResponse
    If InStr(txtResponse, "FAIL") > 0 Then
        SetForeColor vasID, glRow, glRow, 1, colState, 255, 0, 0
        SetText vasID, "확인바람", glRow, colState
        Exit Function
    End If
'    txtXML.Text = txtResponse
'    txtResponse = Replace(txtResponse, ">", ">" & vbNewLine)
'    txtXMLedit.Text = txtResponse
'
    
    
    POCT_XML = True
    Exit Function
err_handler:
  If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description
    
    
End Function

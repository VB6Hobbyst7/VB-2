VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   " EPOC eDM Lite Interface "
   ClientHeight    =   10680
   ClientLeft      =   330
   ClientTop       =   825
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInterface.frx":030A
   ScaleHeight     =   10680
   ScaleWidth      =   15165
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   2655
      Left            =   420
      TabIndex        =   29
      Top             =   7050
      Visible         =   0   'False
      Width           =   8175
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   5640
         Top             =   2070
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   7200
         TabIndex        =   49
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
         SpreadDesigner  =   "frmInterface.frx":058D
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   47
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
         TabIndex        =   36
         Top             =   1380
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   120
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   735
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   33
         Top             =   1320
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
         TabIndex        =   32
         Top             =   1320
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   960
         Left            =   6705
         TabIndex        =   31
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
         TabIndex        =   30
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
         SpreadDesigner  =   "frmInterface.frx":07A5
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   3195
         TabIndex        =   37
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
         SpreadDesigner  =   "frmInterface.frx":09BD
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   4980
         TabIndex        =   38
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
         SpreadDesigner  =   "frmInterface.frx":0BD5
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   1800
         TabIndex        =   39
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
         SpreadDesigner  =   "frmInterface.frx":0DED
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   4860
         TabIndex        =   41
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   5700
         TabIndex        =   40
         Top             =   1410
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   2085
      Left            =   15165
      TabIndex        =   26
      Top             =   3915
      Visible         =   0   'False
      Width           =   9465
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   1260
         TabIndex        =   27
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
         SpreadDesigner  =   "frmInterface.frx":1005
      End
      Begin FPSpread.vaSpread vasPrintBuf 
         Height          =   1245
         Left            =   120
         TabIndex        =   28
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
         SpreadDesigner  =   "frmInterface.frx":2A7E
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
      TabPicture(0)   =   "frmInterface.frx":2C96
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":2CB2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
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
            TabIndex        =   42
            Top             =   630
            Width           =   6015
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   48
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Left            =   4200
               TabIndex        =   46
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
               TabIndex        =   45
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Left            =   1605
               TabIndex        =   44
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
               TabIndex        =   43
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
            TabIndex        =   25
            Top             =   -120
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
            Left            =   2820
            TabIndex        =   24
            Top             =   240
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   345
            Left            =   180
            TabIndex        =   23
            Top             =   270
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   609
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
            Format          =   66846720
            CurrentDate     =   40457
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   690
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
            Left            =   11610
            TabIndex        =   19
            Top             =   240
            Visible         =   0   'False
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
            Left            =   13050
            TabIndex        =   18
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRID1 
            Height          =   7815
            Left            =   9480
            TabIndex        =   21
            Top             =   2370
            Visible         =   0   'False
            Width           =   8055
            _Version        =   393216
            _ExtentX        =   14208
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
            MaxCols         =   12
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":2CCE
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   7275
            Left            =   8460
            TabIndex        =   22
            Top             =   1260
            Width           =   5985
            _Version        =   393216
            _ExtentX        =   10557
            _ExtentY        =   12832
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
            MaxCols         =   7
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":36DE
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   7815
            Left            =   150
            TabIndex        =   55
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
            MaxCols         =   12
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":7479
            UserResize      =   2
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8775
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   14625
         Begin VB.TextBox Text2 
            Height          =   345
            Left            =   7770
            TabIndex        =   57
            Text            =   "EPOC"
            Top             =   270
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            Height          =   345
            Left            =   6480
            TabIndex        =   56
            Text            =   "1954407131"
            Top             =   270
            Width           =   1275
         End
         Begin VB.FileListBox FileEDMLite 
            Height          =   285
            Left            =   5070
            TabIndex        =   54
            Top             =   240
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.CommandButton Command1 
            Caption         =   "TEST"
            Height          =   375
            Left            =   9090
            TabIndex        =   53
            Top             =   240
            Width           =   1185
         End
         Begin VB.Timer tmrResult 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   4140
            Top             =   450
         End
         Begin VB.CommandButton cmdReadEpoc 
            Caption         =   "결과가져오기"
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
            Left            =   150
            TabIndex        =   52
            Top             =   270
            Width           =   1395
         End
         Begin VB.TextBox txtTimer 
            Alignment       =   2  '가운데 맞춤
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1590
            TabIndex        =   51
            Text            =   "20"
            Top             =   270
            Width           =   795
         End
         Begin VB.CheckBox chkSearch 
            Caption         =   "조회중"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2580
            TabIndex        =   50
            Top             =   270
            Value           =   1  '확인
            Width           =   1305
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Command16"
            Height          =   435
            Left            =   6060
            TabIndex        =   12
            Top             =   4950
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTest 
            Height          =   675
            Left            =   1680
            TabIndex        =   11
            Top             =   4800
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
            Left            =   150
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
            MaxCols         =   12
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":7F0F
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
            MaxCols         =   7
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":89A5
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
            Object.Width           =   7223
            MinWidth        =   7233
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5467
            MinWidth        =   5467
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2019-05-20"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 12:33"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "EPOC eDM Lite 실시간 조회"
            TextSave        =   "EPOC eDM Lite 실시간 조회"
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
      Caption         =   "     EPOC eDM Lite INTERFACE"
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
         Left            =   4785
         Picture         =   "frmInterface.frx":C6F0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   195
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
         Format          =   66846720
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
         Left            =   5190
         TabIndex        =   4
         Top             =   255
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
Const colA1c = 13
Const colIFCC = 15
Const coleAg = 17

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
Dim gsPosNo As String
Dim gsResDateTime As String
Dim gsSeqNo As String
Dim gsExamCode As String
Dim gsExamName As String
Dim gsOrder As String
Dim gsResult As String
Dim gsFlag As String

Dim gMT As String
Dim gComState As Long
Dim gErrState As Long

Dim gIFCC1 As String
Dim gIFCC2 As String
Dim geAg1 As String
Dim geAg2 As String
Dim gADD_IFCC As String
Dim gADD_eAg As String

Dim strBuffer As String

Public gENQFlag As Integer
Public gNAKFlag As Integer

'===============================
Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const rs  As String = ""
Const GS  As String = ""

Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer
'===============================

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

Private Sub chkSearch_Click()
    If chkSearch.Value = "1" Then
        tmrResult.Enabled = True
        chkSearch.Caption = "조회중"
    Else
        tmrResult.Enabled = False
        chkSearch.Caption = "대기중"
    End If
    
End Sub

Private Sub cmdExcel_Click()
    Dim iRow As Integer
    Dim j As Integer
    
    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
    Dim sFileName As String
    
    Dim sA1c As String
    Dim sIFCC As String
    Dim seAg As String
    
    ClearSpread vasPrint

    j = 1

    For iRow = 1 To vasRID.DataRowCnt
        vasRID.Row = iRow
        vasRID.Col = 1

        If vasRID.Value = 1 Then
            SetText vasPrint, Trim(GetText(vasRID, iRow, colSpecNo)), j, 1
            SetText vasPrint, Trim(GetText(vasRID, iRow, colBarcode)), j, 2
            SetText vasPrint, Trim(GetText(vasRID, iRow, colPID)), j, 3
            SetText vasPrint, Trim(GetText(vasRID, iRow, colPName)), j, 4
            SetText vasPrint, Trim(GetText(vasRID, iRow, colSex)), j, 5
            'SetText vasPrint, Trim(GetText(vasRID, iRow, colHospital)), j, 5
            
            SQL = "SELECT RESULT " & vbCrLf & _
                  "FROM PAT_RES " & vbCrLf & _
                  "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                  "  AND BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf & _
                  "  AND PID = '" & Trim(GetText(vasPrint, iRow, 3)) & "' " & vbCrLf & _
                  "ORDER BY SEQNO"
            res = db_select_Vas(gLocal, SQL, vasPrintBuf)
            
            sA1c = GetText(vasPrintBuf, 1, 1)
            sIFCC = GetText(vasPrintBuf, 2, 1)
            seAg = GetText(vasPrintBuf, 3, 1)

            ClearSpread vasPrintBuf, 1, 1

            SetText vasPrint, sA1c, j, 7
            SetText vasPrint, sIFCC, j, 8
            SetText vasPrint, seAg, j, 9
            
            j = j + 1
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "저장할 자료가 없습니다.", , "알 림"
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        SaveExcel sFileName, vasPrint
        
    End If
End Sub
Sub SaveExcel(Filename As String, argSpread As vaSpread)

On Error Resume Next

' Excel Object Library 와 연결합니다.
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer

    Set xlApp = CreateObject("Excel.Application")
    
    xlApp.DisplayAlerts = False
    
    Set xlBook = xlApp.Workbooks.Add
    
    Set xlSheet = xlBook.Worksheets(1)
     
    For iRow = 0 To argSpread.DataRowCnt
        For iCol = 1 To argSpread.DataColCnt
            argSpread.Row = iRow
            argSpread.Col = iCol
            xlSheet.Cells(iRow + 1, iCol) = argSpread.Text
        Next iCol
    Next iRow
    
    xlBook.SaveAs (Filename)
    xlApp.Quit


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
            If Mid(Trim(GetText(vasID, lRow, 3)), 1, 2) = "99" Then
                res = Insert_Data_QC(gRow)
            Else
                res = Insert_Data(gRow)
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

Private Sub cmdReadEpoc_Click()
    
    txtTimer.Text = "1"
'''
'''    Dim strFolder   As String
'''    Dim strFile     As String
'''    Dim i, J        As Integer
'''
'''    strFolder = GetDirectory
'''
'''    If strFolder <> "" Then
'''        DirEDM.Path = strFolder
'''        DirEDM.Refresh
'''
'''        DoEvents
'''
'''        For i = 0 To DirEDM.ListCount - 1
'''            DirEDM.ListIndex = i
'''            FileEDM.Path = DirEDM.List(i)
'''            For J = 0 To FileEDM.ListCount - 1
'''                strFile = FileEDM.List(J)
'''                If UCase(strFile) = "INFO.XML" Then
'''
''''                    Call DisplayNode_Info(FileEDM.Path & "\" & strFile)
''''                    Call DisplayNode_Result(FileEDM.Path & "\result.xml")
'''
'''
''''                    If UBound(strRecvData) > 1 Then
''''                        Call SerialRcvData_HPV
''''                    End If
'''                End If
'''            Next
'''            'Fileedm.PATH = diredm.p
'''            'FileIT1000.ListIndex = intIdx
'''
'''            'strSrcfile = FileIT1000.PATH & "\" & FileIT1000.Filename  ' 원본 파일 이름을 정의합니다.
'''
'''        Next
'''    End If
    
End Sub


'================================================================================
' Function    : GetDirectory
' DateTime    : 2007-07-25 18:22
' Author      : S.J.LEE ( http://cafe.naver.com/xlsvba )
' Purpose     : 사용자가 지정한 디렉토리의 Path를 문자열로 취득
' Param       : strMsg - 다이알로그에 출력될 메세지
'================================================================================

Private Function GetDirectory(Optional strMSG As String = "Select A Directory") As String
    Dim BI      As BROWSEINFO
    Dim strPath As String
    Dim lngRet  As Long, x As Long
    Dim intP    As Integer

    With BI
        .ulFlags = &H1
        .pidlRoot = 0&
        .lpszTitle = strMSG
        .lParam = 0
        .lpfn = 0
        .pszDisplayName = ""
        .pidlRoot = 0
        .ulFlags = 0
    End With
   
    x = SHBrowseForFolder(BI)
 
    strPath = String(512, Chr(0))
    
    lngRet = SHGetPathFromIDList(ByVal x, ByVal strPath)
    
    If lngRet Then
        GetDirectory = Replace(strPath, Chr(0), "", , , 1)
    Else
        GetDirectory = ""
    End If
 

End Function


Private Sub cmdRSch_Click()
    Dim iRow As Long

    ClearSpread vasRID
    ClearSpread vasRRes
    Call chkRAll_Click
    
    vasRID.RowHeight(-1) = 12
    
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
            res = Insert_Data_R(lRow)
        
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
'    Dim varXML      As Variant
'    Dim strXmlName  As String
'    Dim i As Integer
'
'  '  StatusBar1.Panels(3).Text = ""
'
'    ' gAssayNM.OrderPath
'    CommonDialog1.Filter = "XML(*.*)|*.*"
'    CommonDialog1.Action = 1
'
'    If CommonDialog1.FileTitle = "" Then
'        Exit Sub
'    End If
'
'    strXmlName = Trim(CommonDialog1.Filename)
'
'    'Call f_subSet_XMLWorkList(strXmlName)
'
''    Call EditRcvDataAPEX_Front
'
'    'Call EditRcvDataAPEX
'
'  '  Call SetComment

    Dim oSOAP   As MSSOAPLib30.SoapClient30
    Dim Send    As String
    Dim sParam  As String
    Dim sBcno   As String
    Dim gEquip  As String
    
    sBcno = Text1.Text
    gEquip = Text2.Text
    
    On Error GoTo ErrHandle
    
'    Online_XML_Qry = ""
    
    Set oSOAP = New MSSOAPLib30.SoapClient30
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    oSOAP.MSSoapInit "http://isis.nhimc:8800/service/PoctService?wsdl" 'gServerPath

    Call Save_Raw_Data("[Send SOAP  => " & sBcno & " ]" & gEquip)
    
    Send = oSOAP.registSpcmRcpn(sBcno, gEquip)
    
    Send = oSOAP.PoctService(sBcno, gEquip)
    
    'MsgBox Send
    
    Call Save_Raw_Data("[Recv SOAP => " & sBcno & " ]" & Send)
    
'    Online_XML_Qry = Send
    Set oSOAP = Nothing
    
    DoEvents
    
    Exit Sub
    
ErrHandle:
'    Online_XML_Qry = ""
    
    MsgBox "[Online_XML_Qry]" & Err.Description
    Save_Raw_Data "[Online_XML_Qry]" & Err.Description

    If oSOAP.FaultString <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[SOAP]" & oSOAP.FaultString & vbCrLf & oSOAP.Detail & vbCrLf
    End If
    If Trim(Err.Description) <> "" Then
        Debug.Print Format(Time, "hh:nn:ss") & "[ERROR]" & Err.Description & vbCrLf
    End If

End Sub

Private Sub lblclear_Click()
    lblChangeBar.Caption = ""
    lblBarcode.Caption = ""
    lblChangePID.Caption = ""
    lblPname.Caption = ""
End Sub

Private Sub Command16_Click()
    Dim i As Long
    Dim lsChar As String
    
'일산병원 Dump
strBuffer = ""
strBuffer = strBuffer & ""
strBuffer = strBuffer & "1H|\^&|||ABL800 BASIC^||||||||1|20110727135635" & vbCrLf
strBuffer = strBuffer & "0C" & vbCrLf
strBuffer = strBuffer & "2P|1||1602800010||^|||U||||||^|^^^|^|^||||||||" & vbCrLf
strBuffer = strBuffer & "D6" & vbCrLf
strBuffer = strBuffer & "3O|1||Sample #^51322|^^^Syringe|||||||||||Not specified^|" & vbCrLf
strBuffer = strBuffer & "ED" & vbCrLf
strBuffer = strBuffer & "4R|1|^^^pCO2^M|26.7|mmHg||N||F||Anonymous|20110727135500" & vbCrLf
strBuffer = strBuffer & "9D" & vbCrLf
strBuffer = strBuffer & "5R|2|^^^pH^M|7.516|||N||F|||" & vbCrLf
strBuffer = strBuffer & "43" & vbCrLf
strBuffer = strBuffer & "6R|3|^^^pO2^M|57.1|mmHg||N||F|||" & vbCrLf
strBuffer = strBuffer & "D1" & vbCrLf
strBuffer = strBuffer & "7R|4|^^^Ca++^M|1.08|mmol/L||N||F|||" & vbCrLf
strBuffer = strBuffer & "7F" & vbCrLf
strBuffer = strBuffer & "0R|5|^^^tHb^M|9.0|g/dL||N||F|||" & vbCrLf
strBuffer = strBuffer & "83" & vbCrLf
strBuffer = strBuffer & "1R|6|^^^sO2^M|91.3|%||N||F|||" & vbCrLf
strBuffer = strBuffer & "6E" & vbCrLf
strBuffer = strBuffer & "2R|7|^^^T^I|37.0|Cel||||F|||" & vbCrLf
strBuffer = strBuffer & "6A" & vbCrLf
strBuffer = strBuffer & "3R|8|^^^HCO3-^C|21.5|mmol/L||||F|||" & vbCrLf
strBuffer = strBuffer & "66" & vbCrLf
strBuffer = strBuffer & "4R|9|^^^tCO2(B)^C|19.8|mmol/L||||F|||" & vbCrLf
strBuffer = strBuffer & "03" & vbCrLf
strBuffer = strBuffer & "5R|10|^^^tO2^E|11.6|Vol%||||F|||" & vbCrLf
strBuffer = strBuffer & "74" & vbCrLf
strBuffer = strBuffer & "6R|11|^^^SBE^C|-1.2|mmol/L||||F|||" & vbCrLf
strBuffer = strBuffer & "2B" & vbCrLf
strBuffer = strBuffer & "7R|12|^^^ABE^C|-0.5|mmol/L||||F|||" & vbCrLf
strBuffer = strBuffer & "1D" & vbCrLf
strBuffer = strBuffer & "0L|1|N" & vbCrLf
strBuffer = strBuffer & "03" & vbCrLf
strBuffer = strBuffer & "" & vbCrLf
    
    
strBuffer = ""
strBuffer = strBuffer & ""
strBuffer = strBuffer & "1H|\^&|||ABL825^||||||||1|20111013112917" & vbCrLf
strBuffer = strBuffer & "84" & vbCrLf
strBuffer = strBuffer & "2P|1||1120020741||^|||U||||||^|^^^|^|^||||||||" & vbCrLf
strBuffer = strBuffer & "CC" & vbCrLf
strBuffer = strBuffer & "3O|1||Sample #^46005|^^^Electro|||||||||||Not specified^|" & vbCrLf
strBuffer = strBuffer & "DC" & vbCrLf
strBuffer = strBuffer & "4C|1|I|688|I" & vbCrLf
strBuffer = strBuffer & "F4" & vbCrLf
strBuffer = strBuffer & "5R|1|^^^Cl-^M|111|mmol/L||N||F||Anonymous|20111013102300" & vbCrLf
strBuffer = strBuffer & "A0" & vbCrLf
strBuffer = strBuffer & "6R|2|^^^K+^M|3.0|mmol/L||N||F|||" & vbCrLf
strBuffer = strBuffer & "C2" & vbCrLf
strBuffer = strBuffer & "7R|3|^^^Na+^M|143|mmol/L||N||F|||" & vbCrLf
strBuffer = strBuffer & "2F" & vbCrLf
strBuffer = strBuffer & "0R|4|^^^Ca++^M|1.15|mmol/L||N||F|||" & vbCrLf
strBuffer = strBuffer & "76" & vbCrLf
strBuffer = strBuffer & "1R|5|^^^T^I|37.0|Cel||||F|||" & vbCrLf
strBuffer = strBuffer & "67" & vbCrLf
strBuffer = strBuffer & "2L|1|N" & vbCrLf
strBuffer = strBuffer & "05" & vbCrLf
strBuffer = strBuffer & "" & vbCrLf
    
'strBuffer = ""
'strBuffer = strBuffer & ""
'strBuffer = strBuffer & "1H|\^&|||ABL800 BASIC^||||||||1|20111013135402" & vbCrLf
'strBuffer = strBuffer & "F9" & vbCrLf
'strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||" & vbCrLf
'strBuffer = strBuffer & "6F" & vbCrLf
'strBuffer = strBuffer & "3O|1||QC #^5303||||||||||||S7735^353" & vbCrLf
'strBuffer = strBuffer & "3D" & vbCrLf
'strBuffer = strBuffer & "4R|1|^^^T^I|25.6|Cel||||F|||20111012080000" & vbCrLf
'strBuffer = strBuffer & "19" & vbCrLf
'strBuffer = strBuffer & "5R|2|^^^pCO2^M|65.9|mmHg||||F|||" & vbCrLf
'strBuffer = strBuffer & "CB" & vbCrLf
'strBuffer = strBuffer & "2L|1|N" & vbCrLf
'strBuffer = strBuffer & "05" & vbCrLf
'strBuffer = strBuffer & "" & vbCrLf
    
    
strBuffer = ""
strBuffer = strBuffer & ""
strBuffer = strBuffer & "1H|\^&|||ABL825^||||||||1|2011110308221985" & vbCrLf
strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||6F" & vbCrLf
strBuffer = strBuffer & "3O|1||QC #^5712||||||||||||S7765^26041" & vbCrLf
strBuffer = strBuffer & "4R|1|^^^T^I|24.8|Cel||||F|||2011110308140020" & vbCrLf
strBuffer = strBuffer & "5R|2|^^^pCO2^M|92.1|mmHg||||F|||C3" & vbCrLf
strBuffer = strBuffer & "6R|3|^^^Cl-^M|90|mmol/L||||F|||B3" & vbCrLf
strBuffer = strBuffer & "7R|4|^^^pH^M|6.817|||||F|||FC" & vbCrLf
strBuffer = strBuffer & "0R|5|^^^pO2^M|293|mmHg||||F|||52" & vbCrLf
strBuffer = strBuffer & "1R|6|^^^Ca++^M|1.64|mmol/L||||F|||2F" & vbCrLf
strBuffer = strBuffer & "2R|7|^^^K+^M|6.3|mmol/L||||F|||7B" & vbCrLf
strBuffer = strBuffer & "3R|8|^^^Na+^M|124|mmol/L||||F|||E1" & vbCrLf
strBuffer = strBuffer & "4R|9|^^^tHb^M|2.6|g/dL||||F|||3C" & vbCrLf
strBuffer = strBuffer & "5R|10|^^^sO2^M|5.0|%||||F|||17" & vbCrLf
strBuffer = strBuffer & "6R|11|^^^O2Hb^M|3.6|%||||F|||54" & vbCrLf
strBuffer = strBuffer & "7R|12|^^^COHb^M|9.0|%||||F|||67" & vbCrLf
strBuffer = strBuffer & "0R|13|^^^MetHb^M|19.6|%||||F|||2C" & vbCrLf
strBuffer = strBuffer & "1R|14|^^^B^M|768|mmHg||||F|||DB" & vbCrLf
strBuffer = strBuffer & "2R|15|^^^pH(T)^C|6.817|||||F|||C4" & vbCrLf
strBuffer = strBuffer & "3R|16|^^^pCO2(T)^C|92.0|mmHg||||F|||90" & vbCrLf
strBuffer = strBuffer & "4R|17|^^^pO2(T)^C|292|mmHg||||F|||23" & vbCrLf
strBuffer = strBuffer & "5L|1|N" & vbCrLf
strBuffer = strBuffer & "08" & vbCrLf
strBuffer = strBuffer & "" & vbCrLf
    
strBuffer = ""
strBuffer = strBuffer & ""
strBuffer = strBuffer & "1H|\^&|||ABL800 BASIC^||||||||1|20111124133125" & vbCrLf
strBuffer = strBuffer & "FC"
strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||" & vbCrLf
strBuffer = strBuffer & "6F"
strBuffer = strBuffer & "3O|1||QC #^5448||||||||||||S7765^260" & vbCrLf
strBuffer = strBuffer & "47"
strBuffer = strBuffer & "4R|1|^^^T^I|27.2|Cel||||F|||20111124080000" & vbCrLf
strBuffer = strBuffer & "1B"

strBuffer = strBuffer & "5R|2|^^^pH^M|6.818|||||F|||" & vbCrLf
strBuffer = strBuffer & "F9"
strBuffer = strBuffer & "6R|3|^^^pO2^M|282|mmHg||||F|||" & vbCrLf
strBuffer = strBuffer & "54"
strBuffer = strBuffer & "7R|4|^^^pCO2^M|95.1|mmHg||||F|||" & vbCrLf
strBuffer = strBuffer & "CA"
strBuffer = strBuffer & "0R|5|^^^Ca++^M|1.66|mmol/L||||F|||" & vbCrLf
strBuffer = strBuffer & "2F"

strBuffer = strBuffer & "1R|6|^^^tHb^M|2.6|g/dL||||F|||" & vbCrLf
strBuffer = strBuffer & "36"
strBuffer = strBuffer & "2R|7|^^^sO2^M|5.0|%||||F|||" & vbCrLf
strBuffer = strBuffer & "EA"
strBuffer = strBuffer & "3R|8|^^^B^M|770|mmHg||||F|||" & vbCrLf
strBuffer = strBuffer & "A9"
strBuffer = strBuffer & "4R|9|^^^pH(T)^C|6.815|||||F|||" & vbCrLf
strBuffer = strBuffer & "97"
strBuffer = strBuffer & "5R|10|^^^pCO2(T)^C|96.5|mmHg||||F|||" & vbCrLf
strBuffer = strBuffer & "95"
strBuffer = strBuffer & "6R|11|^^^pO2(T)^C|288|mmHg||||F|||" & vbCrLf
strBuffer = strBuffer & "24"

strBuffer = strBuffer & "7L|1|N" & vbCrLf
strBuffer = strBuffer & "0A" & vbCrLf
strBuffer = strBuffer & "" & vbCrLf


Call MSComm1_OnComm
    
    Exit Sub
    
    
    For i = 1 To Len(txtTest)
        lsChar = Mid(txtTest, i, 1)

        Select Case lsChar
        Case chrSTX
            txtData.Text = lsChar
            
        Case chrETX
            SaveData "[RX]" & txtData.Text & lsChar
            
            URISCAN_PRO txtData  '한 레코드 받으면 처리
            
        Case Else
            txtData.Text = txtData.Text & lsChar
        End Select
    Next i
    
    txtTest = ""

End Sub

Private Sub URISCAN_PRO(asData As String)
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
    
    Me.Left = 0
    Me.Top = 0
    
    Me.Height = 11520
    Me.Width = 15435
    
    cmdIFClear_Click
    cmdRClear_Click
    lblclear_Click
    
    GetSetup
    
'    MSComm1.CommPort = gSetup.gPort
'    MSComm1.RTSEnable = gSetup.gRTSEnable
'    MSComm1.DTREnable = gSetup.gDTREnable
'    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
'
'    If MSComm1.PortOpen = False Then
'        MSComm1.PortOpen = True
'    End If
    
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
            Exit For
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

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    tmrResult.Interval = 1000
    tmrResult.Enabled = True
    
End Sub

Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "  From equipexam " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by  examcode "
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

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '송신할 데이터

    
    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            intSndPhase = 4
            intFrameNo = intFrameNo + 1
            
        Case 3  '## No Order
            
        Case 4  '## Order
            If mOrder.NoOrder = True Then
                '## 접수정보가 없을경우
                strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & _
                            "^" & mOrder.TubePos & "^^SAMPLE^NORMAL|ALL" & _
                            "|R||||||C||||||||||||||Q" & vbCr & ETX
                intSndPhase = 5
            
            Else
                If mOrder.IsSending = False Then   '## 최초 보낼때
                    strOutput = "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & "^" & mOrder.TubePos & _
                                "^^SAMPLE^NORMAL|" & mOrder.Order & "|R||||||N||||||||||||||Q"
                                
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                Else                        '## 남은 문자열이 있을때
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1
        Case 5  '## Termianator
            strOutput = intFrameNo & "L|1" & vbCr & ETX
            intSndPhase = 6
            intFrameNo = intFrameNo + 1
            
        Case 6  '## EOT
            strState = ""
            MSComm1.Output = EOT
            Save_Raw_Data "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    MSComm1.Output = strOutput
    Debug.Print strOutput
    Save_Raw_Data "[Tx]" & strOutput
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열의 CheckSum을 구함
'   인수 :
'       - pMsg : 문자열
'   반환 : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
End Function

'-- 지금날짜와 검사일자 비교한다
Function DateCompare(ByVal FDate As String) As String
    
    DateCompare = FDate
    If FDate <> Format(Now, "yyyymmdd") Then
        DateCompare = Format(Now, "yyyymmdd")
    End If
    
End Function

Private Sub MSComm1_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long
    Dim strDate As String

    strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
    dtpToday.Value = Format(strDate, "####-##-##")
    DoEvents

    Select Case MSComm1.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long
            
            Buffer = MSComm1.Input
            Save_Raw_Data "[Rx]" & Buffer
            lngBufLen = Len(Buffer)
            
'            Debug.Print Buffer
'            Buffer = strBuffer
'            lngBufLen = Len(Buffer)
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case intPhase
                    Case 1      '## Estabilshment Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                intPhase = 2
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
                            Case ACK
                                '-- 단방향 장비임.
                                'If strState = "Q" Then Call SendOrder
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
                            Case STX
                                intBufCnt = 1
                                Erase strRecvData
                                ReDim Preserve strRecvData(intBufCnt)
                            Case ETB
                                blnIsETB = True
                                intBufCnt = intBufCnt + 1
                                ReDim Preserve strRecvData(intBufCnt)
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
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
                        End Select
                    Case 4      '## Termination Phase
                        Select Case BufChar
                            Case STX
                                intPhase = 2
                            Case EOT
                                Call EditRcvData
                                If strState = "Q" Then
                                    intSndPhase = 1
                                    intFrameNo = 1
                                    MSComm1.Output = ENQ
                                    Save_Raw_Data "[Tx]" & ENQ
                                End If
                                intPhase = 1
                        End Select
                End Select
            Next i

        Case comEvSend
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

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    
    intRow = -1
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBarcode)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    Call SetText(vasID, pBarNo, intRow, colBarcode)  '3
    Call SetText(vasID, mOrder.RackNo, intRow, colRack)       '4
    Call SetText(vasID, mOrder.TubePos, intRow, colPos)         '5
    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    
    strItems = GetEquipExamCode_E411(gEquip, pBarNo)

    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = ""
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
    End If
    

End Sub

'-----------------------------------------------------------------------------'
'   기능 :
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim sRet        As String
    
    intRow = -1
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBarcode)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    Call SetText(vasID, pBarNo, intRow, colBarcode)  '3
    Call SetText(vasID, mResult.RackNo, intRow, colRack)       '4
    Call SetText(vasID, mResult.TubePos, intRow, colPos)         '5
    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    
'    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    
    If Mid(pBarNo, 1, 2) = "99" Then
        Call Get_Sample_Info_QC(intRow)
    Else
        
        Call Online_XML_Qry(pBarNo)
    

        Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    End If
    
    
    vasID.RowHeight(-1) = 12
    
    gRow = intRow
    
    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)


End Sub


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarno     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    
    Dim strQCSeq    As String
    Dim strQCBarNo  As String
    
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        strType = Mid$(strRcvBuf, 2, 1)
        If strType = "|" Then
            strType = Mid$(strRcvBuf, 1, 1)
        End If
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
                strBarno = Format$(mGetP(strRcvBuf, 4, "|"), String$(15, "#"))
                mResult.BarNo = strBarno
            '-- 단방향 장비임.
            Case "Q"    '## Request Information
                '## 바코드번호, SEQ, Disk No, Tube Position 조회
'                If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
'                strTemp1 = mGetP(strRcvBuf, 3, "|")
'                strBarno = Trim$(mGetP(strTemp1, 2, "^"))
'
'                mOrder.NoOrder = False
'                mOrder.BarNo = strBarno
'                mOrder.Seq = mGetP(strTemp1, 3, "^")
'                mOrder.RackNo = mGetP(strTemp1, 4, "^")
'                mOrder.TubePos = mGetP(strTemp1, 5, "^")
'
'                Call GetOrder(strBarno)
'                strState = "Q"
            
            Case "O"    '## Order
                strTemp1 = Format$(mGetP(strRcvBuf, 4, "|"), String$(15, "#"))
                strSeq = mGetP(strTemp1, 2, "^")
                mResult.SpcPos = strSeq
                
                strTemp1 = mGetP(strRcvBuf, 16, "|")
                '--2012.06.29 수정 >> 데이터 포맷이 변경됨 (2012.06.02)
                '3O|1||QC #^6501||||||||||||S7755^39445
                '3O|1||QC #^6151||||||||||||QC level 3^S7755C5
                
                'strQCSeq = mGetP(strTemp1, 1, "^")
                strQCSeq = mGetP(strTemp1, 2, "^")
                
                strQCBarNo = ""
                Select Case strQCSeq
                    Case "S7735": strBarno = gQCCode.Basic1
                    Case "S7745": strBarno = gQCCode.Basic2
                    Case "S7755": strBarno = gQCCode.Basic3
                    Case "S7765": strBarno = gQCCode.Basic4
                End Select
                
                mResult.BarNo = strBarno
                
                Call SetPatInfo(strBarno)

            Case "R"    '## Result
                '## 장비기준 검사명, 결과, Abnormal Flag
                strTemp1 = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                strTemp2 = mGetP(strRcvBuf, 4, "|")
                strFlag = mGetP(strRcvBuf, 7, "|")
                
                strIntBase = strTemp1
                '-- QC 온도보정값
                If InStr(strIntBase, "(T)") > 0 Then
                    strIntBase = Replace(strIntBase, "(T)", "")
                End If
                
                strResult = strTemp2
                
                If InStr(strResult, "?") > 0 Then
                    strResult = ""
                    Exit Sub
                End If
                
                If strResult <> "" Then
                          SQL = "Select examcode, examname, seqno "
                    SQL = SQL & "  From equipexam"
                    SQL = SQL & " Where equipno = '" & gEquip & "' "
                    SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                    SQL = SQL & "   and examcode in (" & gOrderExam & ") "
                    res = db_select_Col(gLocal, SQL)
                    
                    '-- 오더 있을 때
                    If res > 0 Then
                        lsExamCode = Trim(gReadBuf(0))
                        lsExamName = Trim(gReadBuf(1))
                        lsSeqNo = Trim(gReadBuf(2))
                                                
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        strResult = SetResult(strResult, strIntBase)
                        lsResult_Buff = strResult
                        
                        SetText vasID, strResult, gRow, colA1c                   '결과
                        SetText vasID, strComm, gRow, colA1c + 1                  'Flag
                        
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strComm, lsResRow, 7                      'Flag
                                                
                        Save_Local_One gRow, lsResRow, "1", lsEquipRes
                        
                        lsResult_Buff = ""
                    Else
                        '--  오더 없을 때
                        
                              SQL = "Select examcode, examname, seqno "
                        SQL = SQL & "  From equipexam"
                        SQL = SQL & " Where equipno = '" & gEquip & "' "
                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                        'SQL = SQL & "   and examcode in (" & gOrderExam & ") "
                        res = db_select_Col(gLocal, SQL)
                        
                        lsExamCode = Trim(gReadBuf(0))
                        lsExamName = Trim(gReadBuf(1))
                        lsSeqNo = Trim(gReadBuf(2))
                                        
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        strResult = SetResult(strResult, strIntBase)
                        lsResult_Buff = strResult
                        
                        SetText vasID, strResult, gRow, colA1c                   '결과
                        SetText vasID, strComm, gRow, colA1c + 1                  'Flag
                        
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strComm, lsResRow, 7                      'Flag
                                 
                        If lsExamCode <> "" Then
                            Save_Local_One gRow, lsResRow, "1", lsEquipRes
                        End If
                        
                        lsResult_Buff = ""
                    End If
                End If
                
                strState = "R"
                
            Case "C"    '## Comment
                '## Abnormal 결과일때 Comment 저장
                If strFlag <> "N" Then
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strComm = mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")

                    '##  인터페이스 결과 컬렉션의 해당 장비기준 검사명이 존재할때만 Comment를 입력 하도록 수정
                    '========================================================================
                          SQL = "Select examcode, examname, seqno "
                    SQL = SQL & "  From equipexam"
                    SQL = SQL & " Where equipno = '" & gEquip & "' "
                    SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                    SQL = SQL & "   and examcode in (" & gOrderExam & ") "
                    res = db_select_Col(gLocal, SQL)

                    If res > 0 Then
                        lsExamCode = Trim(gReadBuf(0))
                        lsExamName = Trim(gReadBuf(1))
                        lsSeqNo = Trim(gReadBuf(2))

                        SetText vasID, strComm, gRow, colA1c + 1                  'Flag
                        SetText vasRes, strComm, lsResRow, 7                      'Flag
                        strComm = ""

                        Save_Local_One gRow, lsResRow, "1", lsEquipRes

                        lsResult_Buff = ""
                    End If
                    '========================================================================
                End If
                
            Case "L"    '## Terminator
                '## DB에 결과저장
                If strState = "R" Then
                    gOrderExam = ""
                    If MnTransAuto.Checked = True Then
                        If Mid(mResult.BarNo, 1, 2) = "99" Then
                            res = Insert_Data_QC(gRow)
                        Else
                            res = Insert_Data(gRow)
                        End If
                        
                        If res = -1 Then
                            SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                            SetText vasID, "Failed", gRow, colState
                        Else
                           
                            SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                            SetText vasID, "Trans", gRow, colState
                            
                            SQL = " Update pat_res Set " & vbCrLf & _
                                  " sendflag = '2' " & vbCrLf & _
                                  " Where equipno = '" & gEquip & "' " & vbCrLf & _
                                  " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                Exit Sub
                            End If
                            
                        End If
                        
                    End If
                
                    SetText vasID, "Result", gRow, colState
                    strState = ""
                End If
        End Select
    Next

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData_HL7()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strTestDate  As String
    Dim strBarno     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    
    Dim strQCSeq    As String
    Dim strQCBarNo  As String
    
On Error GoTo RST

    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        strType = mGetP(strRcvBuf, 1, "|")
        Select Case strType
            Case "MSH"    '## Header
                strTestDate = mGetP(strRcvBuf, 10, "|")
                strTestDate = Mid(strTestDate, 1, 12)
                
            Case "PID"    '## Patient
                strBarno = mGetP(strRcvBuf, 4, "|")
                
                If strBarno = "" Then
                    strBarno = strTestDate
                End If
                
                mResult.BarNo = strBarno
                
                Call SetPatInfo(strBarno)
            
            
            Case "OBR"    '## Order
                'strTemp1 = Format$(mGetP(strRcvBuf, 4, "|"), String$(15, "#"))
                'strSeq = mGetP(strTemp1, 2, "^")
                'mResult.SpcPos = strSeq
                
                'strTemp1 = mGetP(strRcvBuf, 16, "|")
                
                'strQCSeq = mGetP(strTemp1, 1, "^")
'''                strQCSeq = mGetP(strTemp1, 2, "^")
'''
'''                strQCBarNo = ""
'''                Select Case strQCSeq
'''                    Case "S7735": strBarno = gQCCode.Basic1
'''                    Case "S7745": strBarno = gQCCode.Basic2
'''                    Case "S7755": strBarno = gQCCode.Basic3
'''                    Case "S7765": strBarno = gQCCode.Basic4
'''                End Select
                

            Case "OBX"    '## Result
                '## 장비기준 검사명, 결과, Abnormal Flag
                strIntBase = mGetP(strRcvBuf, 4, "|")
                strResult = mGetP(strRcvBuf, 6, "|")
                'strFlag = mGetP(strRcvBuf, 7, "|")
                
                'strIntBase = strTemp1
                
                '-- QC 온도보정값
'                If InStr(strIntBase, "(T)") > 0 Then
'                    strIntBase = Replace(strIntBase, "(T)", "")
'                End If
                
                'strResult = strTemp2
                
                If InStr(strResult, "?") > 0 Then
                    strResult = ""
                    Exit Sub
                End If
                
               ' strResult = Replace(strResult, "<", "")
               ' strResult = Replace(strResult, ">", "")
               ' strResult = Replace(strResult, "=", "")
               ' strResult = Trim(strResult)
                
                
                If strResult <> "" Then
                          SQL = "Select examcode, examname, seqno "
                    SQL = SQL & "  From equipexam"
                    SQL = SQL & " Where equipno = '" & gEquip & "' "
                    SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                    SQL = SQL & "   and examcode in (" & gOrderExam & ") "
                    res = db_select_Col(gLocal, SQL)
                    
                    '-- 오더 있을 때
                    If res > 0 Then
                        lsExamCode = Trim(gReadBuf(0))
                        lsExamName = Trim(gReadBuf(1))
                        lsSeqNo = Trim(gReadBuf(2))
                                                
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        strResult = SetResult(strResult, strIntBase)
                        lsResult_Buff = strResult
                        
                        SetText vasID, strResult, gRow, colA1c                   '결과
                        SetText vasID, strComm, gRow, colA1c + 1                  'Flag
                        
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strComm, lsResRow, 7                      'Flag
                                                
                        Save_Local_One gRow, lsResRow, "1", lsEquipRes
                        
                        lsResult_Buff = ""
                        strState = "R"
                    Else
                        '--  오더 없을 때
                        
                              SQL = "Select examcode, examname, seqno "
                        SQL = SQL & "  From equipexam"
                        SQL = SQL & " Where equipno = '" & gEquip & "' "
                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                        'SQL = SQL & "   and examcode in (" & gOrderExam & ") "
                        res = db_select_Col(gLocal, SQL)
                        
                        lsExamCode = Trim(gReadBuf(0))
                        lsExamName = Trim(gReadBuf(1))
                        lsSeqNo = Trim(gReadBuf(2))
                                        
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.MaxRows < lsResRow Then
                            vasRes.MaxRows = lsResRow
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        lsEquipRes = strResult
                        strResult = SetResult(strResult, strIntBase)
                        lsResult_Buff = strResult
                        
                        SetText vasID, strResult, gRow, colA1c                   '결과
                        SetText vasID, strComm, gRow, colA1c + 1                  'Flag
                        
                        SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                        SetText vasRes, strResult, lsResRow, colResult           '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                        SetText vasRes, strComm, lsResRow, 7                      'Flag
                                 
                        If lsExamCode <> "" Then
                            Save_Local_One gRow, lsResRow, "1", lsEquipRes
                        End If
                        
                        lsResult_Buff = ""
                        If strState = "" Then
                            strState = ""
                        End If
                    End If
                End If
                vasRes.RowHeight(-1) = 12
                'strState = "R"
                
'            Case "C"    '## Comment
'                '## Abnormal 결과일때 Comment 저장
'                If strFlag <> "N" Then
'                    strTemp1 = mGetP(strRcvBuf, 4, "|")
'                    strComm = mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
'
'                    '##  인터페이스 결과 컬렉션의 해당 장비기준 검사명이 존재할때만 Comment를 입력 하도록 수정
'                    '========================================================================
'                          SQL = "Select examcode, examname, seqno "
'                    SQL = SQL & "  From equipexam"
'                    SQL = SQL & " Where equipno = '" & gEquip & "' "
'                    SQL = SQL & "   and equipcode = '" & strIntBase & "' "
'                    SQL = SQL & "   and examcode in (" & gOrderExam & ") "
'                    res = db_select_Col(gLocal, SQL)
'
'                    If res > 0 Then
'                        lsExamCode = Trim(gReadBuf(0))
'                        lsExamName = Trim(gReadBuf(1))
'                        lsSeqNo = Trim(gReadBuf(2))
'
'                        SetText vasID, strComm, gRow, colA1c + 1                  'Flag
'                        SetText vasRes, strComm, lsResRow, 7                      'Flag
'                        strComm = ""
'
'                        Save_Local_One gRow, lsResRow, "1", lsEquipRes
'
'                        lsResult_Buff = ""
'                    End If
'                    '========================================================================
'                End If
                
            Case "L"    '## Terminator
                '## DB에 결과저장
                If strState = "R" Then
                    gOrderExam = ""
                    If MnTransAuto.Checked = True Then
                        If Mid(mResult.BarNo, 1, 2) = "99" Then
                            res = Insert_Data_QC(gRow)
                        Else
                            res = Insert_Data(gRow)
                        End If
                        
                        If res = -1 Then
                            SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                            SetText vasID, "Failed", gRow, colState
                        Else
                           
                            SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                            SetText vasID, "Trans", gRow, colState
                            
                            SQL = " Update pat_res Set " & vbCrLf & _
                                  " sendflag = '2' " & vbCrLf & _
                                  " Where equipno = '" & gEquip & "' " & vbCrLf & _
                                  " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
                            res = SendQuery(gLocal, SQL)
                            If res = -1 Then
                                SaveQuery SQL
                                Exit Sub
                            End If
                            
                        End If
                        
                    End If
                
                    SetText vasID, "Result", gRow, colState
                    strState = ""
                End If
        End Select
    Next

Exit Sub

RST:
    MsgBox Err.Number & vbCrLf & Err.Description

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
    
    SetResult = Trim(asResult)
    
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
    
'    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
'          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
'    res = SendQuery(gLocal, SQL)
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
'          "POSNO, PID, PNAME, " & vbCrLf & _
'          "PSEX, PAGE, " & vbCrLf & _
'          "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
'          "SEQNO, RESULT, EXAMNAME, SENDFLAG, REFFLAG, EQUIPRESULT, RECENO, SAMPLESEQ) " & vbCrLf & _
'          "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, asRow1, colBarcode)) & "', '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
'          "'" & Trim(GetText(vasID, asRow1, colPos)) & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "', " & vbCrLf & _
'          "'" & Trim(GetText(vasID, asRow1, colSex)) & "', " & 0 & ", " & vbCrLf & _
'          "'" & Trim(sExamDate) & "', '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
'          "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', '" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
'          "'" & asSend & "', '" & Trim(GetText(vasRes, asRow2, 7)) & "', '" & Trim(asEquipResult) & "', '" & Trim(GetText(vasID, asRow1, colSpecNo)) & "', " & vbCrLf & _
'          "'" & Trim(GetText(vasID, asRow1, 0)) & "')"
'    res = SendQuery(gLocal, SQL)


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
          "'" & asSend & "', '" & Trim(GetText(vasRes, asRow2, 7)) & "', '" & Trim(asEquipResult) & "', '" & Trim(GetText(vasID, asRow1, colSpecNo)) & "', " & vbCrLf & _
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

Private Sub tmrResult_Timer()
    
    txtTimer.Text = txtTimer.Text - 1
    
    If txtTimer.Text = "0" Then
        txtTimer.Enabled = False
        txtTimer.Text = "20"
        
        Call OpenHL7
    End If
    
End Sub


Private Function OpenHL7() As Boolean
    Dim strFolder   As String
    Dim strFile     As String
    Dim i           As Integer
    Dim j           As Integer
    Dim strBuf      As String
    Dim lngBufLen   As Long
    Dim iCnt        As Integer
    Dim BufChar     As String
    
    
    'strFolder = "C:\Program Files (x86)\Epocal\EDM Lite\hl7out
    strFolder = gHL7Path
    
    If strFolder <> "" Then
        FileEDMLite.Path = strFolder
        FileEDMLite.Refresh
        
'        DoEvents
        
        For j = 0 To FileEDMLite.ListCount - 1
            strFile = FileEDMLite.List(j)
            If UCase(Right(strFile, 3)) = "HL7" Then
                Open strFolder & strFile For Input As #1
        
                strBuffer = ""
                Do While Not EOF(1)
                    strBuf = strBuf & Input(1, #1)
                Loop
        
                Close #1
            
                lngBufLen = Len(strBuf)
                
                For i = 1 To lngBufLen
                    BufChar = Mid$(strBuf, i, 1)
                    If i = 1 Then
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                    End If
                    
                    Select Case BufChar
                        Case vbLf
                        Case vbCr
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
        
                        Case Else
                            If intBufCnt > 0 Then
                                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                            End If
                    End Select
                Next i
                
                If UBound(strRecvData) > 4 Then
                    Call EditRcvData_HL7
                End If
                
                Erase strRecvData
                intBufCnt = 0
                
                Kill strFolder & strFile
                
                'Debug.Print strBuf
            
            End If
        Next
    End If
        
    
    
                                
        
'    If Len(AllergyFile.Filename) > 0 Then
'        Screen.MousePointer = 11
'        Open AllergyFile.Filename For Input As #3
'
'        strBuffer = ""
''        Do While Not EOF(3)
''            strBuf = strBuf & Input(1, #3)
''        Loop
'
'        strBuf = ReadFileBinary(AllergyFile.Filename)
'
'        Close #3
'
'        lngBufLen = Len(strBuf)
'
'
'        For i = 1 To lngBufLen
'            BufChar = Mid$(strBuf, i, 1)
'            Select Case BufChar
'                Case vbLf
'                    iCnt = iCnt + 1
'                    ReDim Preserve strRecvData(iCnt)
'                    strRecvData(iCnt) = strBuffer
'                    strBuffer = ""
'
'                Case Else
''                    If Len(BufChar) > 1 And InStr(BufChar, ",") > 0 Then
''                        BufChar = Replace(BufChar, ",", " ")
''                    End If
'                    strBuffer = strBuffer & BufChar
'
'            End Select
'        Next i
'
'        Call FileRcvData_B4C
'
'    Else
'        Exit Function
'    End If

    Screen.MousePointer = 0
    
    Exit Function

RST:
    Screen.MousePointer = 0

End Function

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

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' " & vbCrLf & _
          " AND DISKNO = '" & Trim(GetText(vasID, Row, colRack)) & "' " & vbCrLf & _
          " AND POSNO = '" & Trim(GetText(vasID, Row, colPos)) & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "
    SQL = SQL & "ORDER BY SEQNO * 10"
    
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.MaxRows = vasRes.DataRowCnt
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
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, REFFLAG, EQUIPRESULT " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          " AND DISKNO = '" & Trim(GetText(vasRID, Row, colRack)) & "' " & vbCrLf & _
          " AND POSNO = '" & Trim(GetText(vasRID, Row, colPos)) & "' " & vbCrLf & _
          " AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, REFFLAG , EQUIPRESULT"
    SQL = SQL & " ORDER BY SEQNO * 10   "
    
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
    
    vasRRes.RowHeight(-1) = 12
    
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

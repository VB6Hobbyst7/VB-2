VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   Caption         =   " URIT8021A Interface"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   15195
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
   LockControls    =   -1  'True
   Picture         =   "frmInterface.frx":030A
   ScaleHeight     =   15390
   ScaleWidth      =   28680
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   3075
      Left            =   4830
      TabIndex        =   13
      Top             =   5940
      Visible         =   0   'False
      Width           =   9435
      Begin FPSpread.vaSpread vasRRRes 
         Height          =   1215
         Left            =   8460
         TabIndex        =   77
         Top             =   360
         Width           =   675
         _Version        =   393216
         _ExtentX        =   1191
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
         SpreadDesigner  =   "frmInterface.frx":058D
      End
      Begin FPSpread.vaSpread vasExamTemp 
         Height          =   1185
         Left            =   5430
         TabIndex        =   72
         Top             =   1410
         Width           =   1335
         _Version        =   393216
         _ExtentX        =   2355
         _ExtentY        =   2090
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
         SpreadDesigner  =   "frmInterface.frx":07B3
      End
      Begin FPSpread.vaSpread vasListTemp 
         Height          =   1455
         Left            =   8340
         TabIndex        =   71
         Top             =   660
         Width           =   495
         _Version        =   393216
         _ExtentX        =   873
         _ExtentY        =   2566
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
         SpreadDesigner  =   "frmInterface.frx":09D9
      End
      Begin FPSpread.vaSpread vasCode_1 
         Height          =   975
         Left            =   6780
         TabIndex        =   68
         Top             =   1740
         Width           =   1635
         _Version        =   393216
         _ExtentX        =   2884
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
         SpreadDesigner  =   "frmInterface.frx":0BFF
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   1275
         Left            =   4920
         TabIndex        =   67
         Top             =   1680
         Width           =   1695
         _Version        =   393216
         _ExtentX        =   2990
         _ExtentY        =   2249
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
         SpreadDesigner  =   "frmInterface.frx":0E25
      End
      Begin FPSpread.vaSpread vasOrderBuf 
         Height          =   1185
         Left            =   2400
         TabIndex        =   66
         Top             =   1620
         Visible         =   0   'False
         Width           =   1875
         _Version        =   393216
         _ExtentX        =   3307
         _ExtentY        =   2090
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
         MaxCols         =   2
         SpreadDesigner  =   "frmInterface.frx":104B
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   975
         Left            =   6780
         TabIndex        =   31
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
         SpreadDesigner  =   "frmInterface.frx":4BA1
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   1815
         Left            =   6840
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         Begin VB.FileListBox FileURIT 
            Height          =   675
            Left            =   0
            Pattern         =   "*.txt"
            TabIndex        =   82
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin MSCommLib.MSComm MSComm1 
            Left            =   540
            Top             =   180
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
         TabIndex        =   23
         Top             =   1320
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   17
         Top             =   1320
         Visible         =   0   'False
         Width           =   1875
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
         TabIndex        =   16
         Top             =   735
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1125
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
         TabIndex        =   14
         Top             =   1380
         Visible         =   0   'False
         Width           =   1635
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   1350
         TabIndex        =   18
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
         SpreadDesigner  =   "frmInterface.frx":4DC7
      End
      Begin FPSpread.vaSpread vasList_1 
         Height          =   1125
         Left            =   3195
         TabIndex        =   19
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
         SpreadDesigner  =   "frmInterface.frx":4FED
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   4980
         TabIndex        =   20
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
         SpreadDesigner  =   "frmInterface.frx":5213
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   1185
         Left            =   60
         TabIndex        =   65
         Top             =   1800
         Visible         =   0   'False
         Width           =   1875
         _Version        =   393216
         _ExtentX        =   3307
         _ExtentY        =   2090
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
         MaxCols         =   2
         SpreadDesigner  =   "frmInterface.frx":5439
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   9315
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   16431
      _Version        =   393216
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
      TabPicture(0)   =   "frmInterface.frx":8F8F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":8FAB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "결과변환"
      TabPicture(2)   =   "frmInterface.frx":8FC7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CommonDialog1"
      Tab(2).Control(1)=   "cmdExcel"
      Tab(2).Control(2)=   "cmdExcelSch"
      Tab(2).Control(3)=   "vasResList"
      Tab(2).Control(4)=   "dtpESDate"
      Tab(2).Control(5)=   "dtpEEDate"
      Tab(2).Control(6)=   "Label12"
      Tab(2).Control(7)=   "Label5"
      Tab(2).ControlCount=   8
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -67080
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel File 변환"
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
         Left            =   -68820
         TabIndex        =   51
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdExcelSch 
         Caption         =   "검사결과조회"
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
         Left            =   -70440
         TabIndex        =   48
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Height          =   8775
         Left            =   -74820
         TabIndex        =   25
         Top             =   360
         Width           =   14625
         Begin VB.CommandButton Command4 
            Caption         =   "검사결과삭제"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   5550
            TabIndex        =   81
            Top             =   8190
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.CommandButton cmdHospDel 
            Caption         =   "검사목록초기화"
            Height          =   495
            Left            =   3840
            TabIndex        =   76
            Top             =   8160
            Width           =   1635
         End
         Begin VB.CommandButton cmdHosp 
            Caption         =   "검사목록저장"
            Height          =   495
            Left            =   2100
            TabIndex        =   75
            Top             =   8160
            Width           =   1635
         End
         Begin VB.ComboBox cmbHosp 
            Height          =   315
            ItemData        =   "frmInterface.frx":8FE3
            Left            =   5640
            List            =   "frmInterface.frx":8FE5
            TabIndex        =   69
            Text            =   "미지정"
            Top             =   300
            Width           =   1695
         End
         Begin VB.CommandButton cmdListMach 
            Caption         =   "LIST 매칭"
            Height          =   495
            Left            =   240
            TabIndex        =   56
            Top             =   8160
            Width           =   1755
         End
         Begin VB.ComboBox cmbServerType 
            Height          =   315
            ItemData        =   "frmInterface.frx":8FE7
            Left            =   3240
            List            =   "frmInterface.frx":8FE9
            TabIndex        =   54
            Top             =   300
            Width           =   1515
         End
         Begin VB.Frame Frame5 
            Height          =   495
            Left            =   7380
            TabIndex        =   44
            Top             =   180
            Width           =   2835
            Begin VB.OptionButton optExamPart 
               Caption         =   "Manual결과"
               Height          =   255
               Index           =   1
               Left            =   1380
               TabIndex        =   46
               Top             =   180
               Width           =   1395
            End
            Begin VB.OptionButton optExamPart 
               Caption         =   "검사결과"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   45
               Top             =   180
               Value           =   -1  'True
               Width           =   1815
            End
         End
         Begin VB.CommandButton cmdRSch 
            Caption         =   "검사결과조회"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10320
            TabIndex        =   30
            Top             =   240
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1020
            TabIndex        =   29
            Top             =   300
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
            Format          =   67960833
            CurrentDate     =   40457
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   660
            TabIndex        =   28
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
            Height          =   435
            Left            =   13080
            TabIndex        =   27
            Top             =   240
            Width           =   1335
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
            Height          =   435
            Left            =   11700
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   7755
            Left            =   8640
            TabIndex        =   42
            Top             =   720
            Width           =   5805
            _Version        =   393216
            _ExtentX        =   10239
            _ExtentY        =   13679
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            EditEnterAction =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   11
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":8FEB
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   7335
            Left            =   240
            TabIndex        =   43
            Top             =   720
            Width           =   8175
            _Version        =   393216
            _ExtentX        =   14420
            _ExtentY        =   12938
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
            MaxCols         =   15
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":CE98
            UserResize      =   2
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "검사구분"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   4860
            TabIndex        =   70
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "검사선택"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   2460
            TabIndex        =   55
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "결과일자"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8775
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   14625
         Begin VB.CommandButton Command3 
            Caption         =   "검사결과조회"
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
            Left            =   12870
            TabIndex        =   79
            Top             =   330
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   540
            TabIndex        =   78
            Top             =   1020
            Width           =   225
         End
         Begin VB.TextBox txtBuff 
            Height          =   1095
            Left            =   5040
            TabIndex        =   64
            Top             =   6840
            Visible         =   0   'False
            Width           =   3195
         End
         Begin Threed.SSCommand cmdOrderTrans 
            Height          =   435
            Left            =   120
            TabIndex        =   52
            Top             =   8220
            Visible         =   0   'False
            Width           =   4185
            _Version        =   65536
            _ExtentX        =   7382
            _ExtentY        =   767
            _StockProps     =   78
            Caption         =   ">>>>>"
         End
         Begin VB.Frame Frame4 
            Caption         =   "[응급여부]"
            Height          =   675
            Left            =   120
            TabIndex        =   39
            Top             =   8580
            Visible         =   0   'False
            Width           =   3855
            Begin VB.TextBox txtRack 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               Enabled         =   0   'False
               Height          =   285
               Left            =   1680
               TabIndex        =   62
               Text            =   "1"
               Top             =   240
               Width           =   825
            End
            Begin VB.TextBox txtPos 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               Enabled         =   0   'False
               Height          =   285
               Left            =   2790
               TabIndex        =   61
               Text            =   "1"
               Top             =   240
               Width           =   645
            End
            Begin VB.CheckBox chkEM 
               Height          =   315
               Left            =   1320
               TabIndex        =   59
               Top             =   240
               Width           =   195
            End
            Begin VB.TextBox txtStartNo 
               Alignment       =   2  '가운데 맞춤
               Appearance      =   0  '평면
               Height          =   285
               Left            =   1680
               TabIndex        =   41
               Text            =   "1"
               Top             =   480
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label Label10 
               Caption         =   "-"
               Height          =   195
               Left            =   2580
               TabIndex        =   63
               Top             =   300
               Width           =   195
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "※ 응급검사"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   180
               Left            =   240
               TabIndex        =   60
               Top             =   300
               Width           =   960
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "※ Sample No."
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Left            =   240
               TabIndex        =   40
               Top             =   540
               Visible         =   0   'False
               Width           =   1245
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "[접수내역조회]"
            Height          =   675
            Left            =   120
            TabIndex        =   33
            Top             =   180
            Width           =   8895
            Begin VB.CommandButton cmdWorklist 
               Caption         =   "접수목록조회"
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
               Left            =   7290
               TabIndex        =   36
               Top             =   210
               Width           =   1395
            End
            Begin VB.CheckBox chkWork 
               Caption         =   "오더전송환자"
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
               Left            =   6690
               TabIndex        =   80
               Top             =   240
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.ComboBox cmbPart 
               Height          =   315
               ItemData        =   "frmInterface.frx":D9C2
               Left            =   5160
               List            =   "frmInterface.frx":D9C4
               TabIndex        =   38
               Top             =   240
               Width           =   1395
            End
            Begin MSComCtl2.DTPicker dtpReceDate 
               Height          =   315
               Left            =   990
               TabIndex        =   34
               Top             =   240
               Width           =   1395
               _ExtentX        =   2461
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
               Format          =   67960833
               CurrentDate     =   40457
            End
            Begin MSComCtl2.DTPicker dtpEndDate 
               Height          =   315
               Left            =   2760
               TabIndex        =   57
               Top             =   240
               Width           =   1395
               _ExtentX        =   2461
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
               Format          =   67960833
               CurrentDate     =   40457
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "~"
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
               Left            =   2460
               TabIndex        =   58
               Top             =   270
               Width           =   180
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "검사선택"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Left            =   4290
               TabIndex        =   37
               Top             =   300
               Width           =   720
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "접수일자"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   180
               Left            =   180
               TabIndex        =   35
               Top             =   300
               Width           =   720
            End
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Command16"
            Height          =   435
            Left            =   8280
            TabIndex        =   10
            Top             =   8280
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTest 
            Height          =   675
            Left            =   4020
            TabIndex        =   9
            Top             =   8040
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
            Left            =   10560
            TabIndex        =   22
            Top             =   360
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
            Left            =   9120
            TabIndex        =   21
            Top             =   360
            Width           =   1395
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   4890
            TabIndex        =   8
            Top             =   1080
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   7725
            Left            =   4500
            TabIndex        =   12
            Top             =   960
            Width           =   5235
            _Version        =   393216
            _ExtentX        =   9234
            _ExtentY        =   13626
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
            MaxCols         =   16
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":D9C6
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   7755
            Left            =   9900
            TabIndex        =   11
            Top             =   960
            Width           =   4545
            _Version        =   393216
            _ExtentX        =   8017
            _ExtentY        =   13679
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
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":E515
         End
         Begin FPSpread.vaSpread vasList 
            Height          =   7755
            Left            =   120
            TabIndex        =   32
            Top             =   960
            Width           =   4215
            _Version        =   393216
            _ExtentX        =   7435
            _ExtentY        =   13679
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
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
            MaxCols         =   10
            MaxRows         =   20
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":1234D
         End
      End
      Begin FPSpread.vaSpread vasResList 
         Height          =   8145
         Left            =   -74880
         TabIndex        =   47
         Top             =   900
         Width           =   14625
         _Version        =   393216
         _ExtentX        =   25797
         _ExtentY        =   14367
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         EditEnterAction =   5
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
         GridColor       =   16777215
         MaxCols         =   50
         MaxRows         =   501
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":12D6A
      End
      Begin MSComCtl2.DTPicker dtpESDate 
         Height          =   315
         Left            =   -73800
         TabIndex        =   49
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   67960833
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker dtpEEDate 
         Height          =   315
         Left            =   -72060
         TabIndex        =   73
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   67960833
         CurrentDate     =   40457
      End
      Begin VB.Label Label12 
         Caption         =   "-"
         Height          =   315
         Left            =   -72300
         TabIndex        =   74
         Top             =   540
         Width           =   135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   -74700
         TabIndex        =   50
         Top             =   540
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   15015
      Width           =   28680
      _ExtentX        =   50588
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
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
            TextSave        =   "2013-03-11"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오전 9:55"
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
      Width           =   15075
      _Version        =   65536
      _ExtentX        =   26591
      _ExtentY        =   1138
      _StockProps     =   15
      Caption         =   "     URIT 8021A INTERFACE"
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
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   10320
         Top             =   150
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4020
         Picture         =   "frmInterface.frx":1771C
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   3
         Top             =   150
         Visible         =   0   'False
         Width           =   315
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
         Format          =   67960832
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
         Caption         =   "사용자"
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
         Left            =   4470
         TabIndex        =   4
         Top             =   210
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
   Begin VB.Menu Mnlist 
      Caption         =   "리스트삭제"
      Visible         =   0   'False
      Begin VB.Menu MnvasListDel 
         Caption         =   "리스트삭제"
      End
   End
   Begin VB.Menu Mnlist2 
      Caption         =   "리스트삭제"
      Visible         =   0   'False
      Begin VB.Menu MnvasIdDel 
         Caption         =   "리스트삭제"
      End
   End
   Begin VB.Menu Mnlist3 
      Caption         =   "리스트삭제"
      Visible         =   0   'False
      Begin VB.Menu MnvasRIDDel 
         Caption         =   "리스트삭제"
      End
      Begin VB.Menu subdel 
         Caption         =   "데이터삭제"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "전송"
      Begin VB.Menu MnTransAuto 
         Caption         =   "자동"
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "수동"
         Checked         =   -1  'True
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
Const colCheckBox = 1
Const colReceDate = 2
Const colSampleNo = 3
Const colRack = 4
Const colPos = 5
Const colReceno = 6
Const colPID = 7
Const colPName = 8
Const colPAge = 9
Const colPSex = 10
Const colPJumin = 11
Const colOCnt = 12
Const colRCnt = 13
Const colState = 14
Const colExamDate = 15
Const colExamGubun = 16

'recedate, sampleno, diskno, posno, receno, pid, pname, page, psex, pjumin, O, R, snedflag
'sendflag
'0: Result 결과만 나온경우
'1: 검사리스트 작성시
'2: 검사리스트 작성후 결과 입력시, Result 결과와 Worklist 매칭시
'3: 결과 전송시

'vasres, vasrres colum
Const colEquipCode = 1
Const colExamCode = 2
Const colSubCode = 3
Const colExamName = 4
Const colResult = 5
Const colSeq = 6
Const colRefFlag = 7


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

Dim gMT As String
Dim gComState As Long
Dim gErrState As Long

Dim llRow As Integer

Private Sub Check1_Click()
    Dim iRow As Long
    
    If Check1.Value = 1 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            vasList.Value = 1
        Next iRow
    ElseIf Check1.Value = 0 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            vasList.Value = 0
        Next iRow
    End If
End Sub

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

Private Sub chkEM_Click()
    If chkEM.Value = 1 Then
        txtRack.Enabled = True
        txtPos.Enabled = True
    ElseIf chkEM.Value = 0 Then
        txtRack.Enabled = True
        txtPos.Enabled = True
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

Private Sub cmdExcel_Click()
    Dim sFileName As String
    Dim xlApp1 As Excel.Application
    Dim xlBook1 As Excel.Workbook
    Dim xlSheet1 As Excel.Worksheet
    Dim lRow, lCol As Long
    Dim sExcelStr As String
    Dim FilNum
    Dim i As Long
    Dim j As Long
    
        
    If vasResList.DataRowCnt <= 0 Or vasResList.DataColCnt <= 0 Then
        MsgBox "저장할 데이터가 없습니다.", vbOKOnly, "Excel"
        Exit Sub
    End If
    
    CommonDialog1.Filter = "Excel Files (*.csv)|*.csv|All Files (*.*)|*.*"
    CommonDialog1.FileName = ""
    CommonDialog1.ShowSave
    sFileName = CommonDialog1.FileName
    If Trim(sFileName) = "" Then
        Exit Sub
    End If
    
    
    'argSQL의 내용을 파일로 저장
    
    sExcelStr = ""
    
    For j = 0 To vasResList.DataRowCnt
        For i = 2 To vasResList.DataColCnt
            If i = 2 And j = 0 Then
                sExcelStr = Trim(GetText(vasResList, j, i))
            ElseIf i = 2 And j <> 0 Then
                sExcelStr = sExcelStr & vbCrLf & Trim(GetText(vasResList, j, i))
            Else
                sExcelStr = sExcelStr & "," & Trim(GetText(vasResList, j, i))
            End If
        Next
    Next
    
    
    FilNum = FreeFile
    
    Open sFileName For Append As FilNum
    Print #FilNum, sExcelStr
    Close FilNum
    
    Exit Sub
End Sub

Private Sub cmdExcelSch_Click()
      Dim i, ii, j, k As Integer

    ClearSpread vasTemp
    
    SQL = "Select distinct equipcode, '', examname, seqno " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "group by  seqno, equipcode, examname "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    For i = 1 To vasTemp.DataRowCnt

        vasResList.MaxCols = 8 + vasTemp.DataRowCnt
        vasResList.SetText 8 + i, 0, Trim(GetText(vasTemp, i, 3))
        vasResList.ColWidth(8 + i) = 8

    Next i
    
    ClearSpread vasResList

    SQL = "SELECT '', examdate, recedate, receno, sampleno, pid, pname " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE between '" & Format(dtpESDate, "YYYYMMDD") & "' and '" & Format(dtpEEDate, "YYYYMMDD") & "' " & vbCrLf & _
          "AND EQUIPNO = '" & gEquip & "' and sendflag <> '1'"
    SQL = SQL & "GROUP BY examdate, recedate, sampleno,  receno, pid, pname"
    res = db_select_Vas(gLocal, SQL, vasResList)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For ii = 1 To vasResList.DataRowCnt
        ClearSpread vasRRRes
        
        SQL = "select '', equipcode,'','',examname,result" & vbCrLf & _
              "from pat_res " & vbCrLf & _
              "where equipno = '" & gEquip & "' " & vbCrLf & _
              "and examdate = '" & GetText(vasResList, ii, 2) & "' " & vbCrLf & _
              "and isnull(receno, '') = '" & GetText(vasResList, ii, 4) & "' " & vbCrLf & _
              "AND isnull(PID, '') = '" & GetText(vasResList, ii, 6) & "'" & vbCrLf & _
              "AND sampleno = '" & GetText(vasResList, ii, 5) & "' and sendflag <> '1' " & vbCrLf & _
              "group by equipcode,examname,result "
        res = db_select_Vas(gLocal, SQL, vasRRRes)
        
     For j = 1 To vasRRRes.DataRowCnt
            For k = 1 To UBound(gArrEquip)
                If Trim(gArrEquip(k, 2)) = Trim(GetText(vasRRRes, j, 2)) Then
                    vasResList.SetText 8 + k, ii, Trim(GetText(vasRRRes, j, 6))
                    Exit For
                End If
            Next k
     Next j
    Next ii
End Sub

Private Sub cmdHosp_Click()
    Dim i As Long
    Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim x As Long
    Dim j As Long
    Dim c, r, c2, r2
    Dim strValue As String
    
    
    strValue = InputBox("검사목록명을 지정하세요.", "검사목록지정")
    
    If Trim(strValue) = "" Then
        Exit Sub
    End If
    
    For i = 1 To vasRID.DataRowCnt
        vasRID.Col = 1
        vasRID.Row = i
        If vasRID.Value = 1 Then
            SQL = "update pat_res set hospital = '" & strValue & "' " & vbCrLf & _
                  "where examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' " & vbCrLf & _
                  "and sampleno = '" & Trim(GetText(vasRID, i, colSampleNo)) & "' and receno = '" & Trim(GetText(vasRID, i, colReceno)) & "' and hospital = '" & cmbHosp.Text & "'" & vbCrLf & _
                  "and equipno = '" & gEquip & "' "
            res = SendQuery(gLocal, SQL)
        End If
        
    Next

'''    If vasRID.IsBlockSelected Or vasRID.SelectionCount Then
'''
'''        vasRID.BlockMode = True
''''        db_BeginTran gLocal
'''
'''        For x = 0 To vasRID.SelectionCount - 1
'''            vasRID.GetSelection x, c, r, c2, r2
'''            vasRID.Col = c
'''            vasRID.Col2 = c2
'''            vasRID.Row = r
'''            vasRID.Row2 = r2
'''            If IsNumeric(r) = True And IsNumeric(r2) = True Then
'''                If CInt(r) > 0 And CInt(r2) > 0 Then
'''                    For i = r To r2
'''                        SQL = "update pat_res set hospital = '" & strValue & "' " & vbCrLf & _
'''                              "where examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' " & vbCrLf & _
'''                              "and sampleno = '" & Trim(GetText(vasRID, i, colSampleNo)) & "' and hospital = '" & cmbHosp.Text & "'"
'''                        res = SendQuery(gLocal, SQL)
'''
'''                    Next
'''                End If
'''            End If
'''        Next x
'''        vasRID.BlockMode = False
'''
'''    End If
    HospChk Format(dtpExamDate, "yyyymmdd")
    cmdRSch_Click
End Sub

Private Sub cmdHospDel_Click()
    Dim i As Long
    Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim x As Long
    Dim j As Long
    Dim c, r, c2, r2
    Dim strValue As String
    Dim msgRes
    
    
    msgRes = MsgBox("목록을 초기화 하시겠습니까?", vbYesNo, "목록초기화")
    If msgRes = 7 Then
        Exit Sub
    End If
    
    For i = 1 To vasRID.DataRowCnt
        vasRID.Col = 1
        vasRID.Row = i
        If vasRID.Value = 1 Then
    
            SQL = "update pat_res set recedate = '', receno = '', sendflag = '0', pid = '', pname = '', page = '', psex = '', pjumin = '', examcode = '', subcode = '', examgubun = '' , hospital = '미지정' " & vbCrLf & _
                  "where examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' " & vbCrLf & _
                  "and sampleno = '" & Trim(GetText(vasRID, i, colSampleNo)) & "' and hospital = '" & cmbHosp.Text & "'" & vbCrLf & _
                  "and equipno = '" & gEquip & "' "
            res = SendQuery(gLocal, SQL)
        End If
        
        
    Next
                    
    HospChk Format(dtpExamDate, "yyyymmdd")
    
    cmdRSch_Click
    
End Sub

Private Sub cmdIFClear_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    ClearSpread vasList
    
    vasRes.MaxRows = 0
    
    gRow = 0
    
    For i = vasID.DataRowCnt To 1 Step -1
        vasID.Row = i
        vasID.Col = 1
        If vasID.Value = 1 Then
            DeleteRow vasID, i, i
        End If
        
    Next
    vasID.MaxRows = vasID.DataRowCnt
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            res = Insert_Data(lRow)
        
            If res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "실패", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "완료", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " SENDFLAG = '3' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND receno = '" & Trim(GetText(vasID, lRow, colReceno)) & "' " & vbCrLf & _
                      " AND sampleno = '" & Trim(GetText(vasID, lRow, colSampleNo)) & "' " & vbCrLf & _
                      " AND examdate = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
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

Private Sub cmdListMach_Click()

    If frmList.Visible = True Then
        MsgBox "이미 창이 존재하고 있습니다!", vbExclamation
        Exit Sub
    End If
    
    frmList.lblExamDate.Caption = Format(dtpExamDate, "yyyy-mm-dd")
    
    frmList.Show 0
    
End Sub

Private Sub cmdOrderTrans_Click()
    Dim i As Long
    Dim j As Long
    Dim x As Long
    Dim intListRow As Long
    Dim intSampleNo As Long
    Dim chServer
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSubCode As String
    Dim lsSeqNo As String
    Dim lsEquipCode As String
    'Order 만들고 전송하기
    Dim sRetOrder As String     'Order Text넣을 변수
    Dim sOrder As String
    
    Dim iRow As Integer

    
    Dim llRow As Long
    Dim llRow_Order As Long
     
    Dim sBarcode As String      '검체번호
    Dim sPID As String
    Dim sReceNo As String       '접수번호
    Dim sRackNo As String
    Dim sPosNo As String
    Dim sORDT As String         '접수일자
    Dim sExamCode As String     '검사코드
    Dim sEquipCode As String    '장비코드
    Dim sOrderCode As String
    Dim sState As String
    
    
    Dim lsCurDate As String
    Dim lsSampleNo As String
    Dim lsType As String
    
    Dim S  As String
    Dim k As Integer
    Dim ii As Long
    Dim yy As Long
    Dim sCnt As String
    Dim lsEquipCodeYN As Integer
    Dim result As String
        
        
    
'''    chServer = cmbPart.ListIndex
    
    For i = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = i
        If vasList.Value = 1 Then
            intListRow = -1
            
            For j = 1 To vasID.DataRowCnt
                If Trim(GetText(vasID, j, colSampleNo)) <> "" And Trim(GetText(vasID, j, colReceno)) = "" Then
                    intListRow = j
                    Exit For
                End If
            Next j
            
            If intListRow = -1 Then
                intListRow = vasID.DataRowCnt + 1
            End If
            
            If intListRow > vasID.MaxRows Then
                vasID.MaxRows = intListRow
            End If
'            intSampleNo = Trim(GetText(vasList, i, 3))
'            txtStartNo = intSampleNo + 1
    
            SetText vasID, Trim(GetText(vasList, i, 2)), intListRow, colReceDate
'''            SetText vasID, Trim(GetText(vasList, i, 3)), intListRow, colSampleNo
            SetText vasID, Trim(GetText(vasList, i, 4)), intListRow, colReceno
            SetText vasID, Trim(GetText(vasList, i, 5)), intListRow, colPID
            SetText vasID, Trim(GetText(vasList, i, 6)), intListRow, colPName
            SetText vasID, Trim(GetText(vasList, i, 7)), intListRow, colPSex
            SetText vasID, Trim(GetText(vasList, i, 8)), intListRow, colPAge
            SetText vasID, Trim(GetText(vasList, i, 9)), intListRow, colPJumin
            SetText vasID, Format(Date, "yyyymmdd"), intListRow, colExamDate
            SetText vasID, Trim(GetText(vasList, i, 10)), intListRow, colExamGubun
            
 '           SetText vasID, CStr(intSampleNo), intListRow, colSampleNo
            chServer = Trim(GetText(vasList, i, 10))
                    
'''            SQL = "select sampleno from pat_res " & vbCrLf & _
'''                  "where recedate = '" & Trim(GetText(vasList, i, 2)) & "' and pid = '" & Trim(GetText(vasList, i, 5)) & "' " & vbCrLf & _
'''                  "and receno = '" & Trim(GetText(vasList, i, 4)) & "' and examgubun = '" & Trim(GetText(vasList, i, 10)) & "' "
'''            res = db_select_Col(gLocal, SQL)
'''            If res > 0 Then
'''            Else
'''
            ClearSpread vasTemp
            
            Select Case chServer
            Case dpGumjin1
                SQL = "select exam_code, '' from totres " & vbCrLf & _
                      "where request_date = '" & Trim(GetText(vasID, intListRow, colReceDate)) & "' " & vbCrLf & _
                      "and exam_no = '" & Trim(GetText(vasID, intListRow, colReceno)) & "' and exam_code in (" & gAllExam & ") "
                res = db_select_Vas(gServer, SQL, vasTemp)
                
            Case dpOCS
                SQL = "select 검사코드, 검사종류 from tb_진료검사 " & vbCrLf & _
                      "where 년 = '" & Mid(Trim(GetText(vasID, intListRow, colReceDate)), 1, 4) & "' " & vbCrLf & _
                      "and 월 = '" & Mid(Trim(GetText(vasID, intListRow, colReceDate)), 5, 2) & "' " & vbCrLf & _
                      "and 일 = '" & Mid(Trim(GetText(vasID, intListRow, colReceDate)), 7, 2) & "' and 챠트번호 = '" & Trim(GetText(vasID, intListRow, colReceno)) & "' and 검사코드 in (" & gAllExam_Ocs & ") and 오더일련번호 > '0' "
                res = db_select_Vas(gServer_OCS, SQL, vasTemp)
            Case dpGumjin2
                SQL = "select exam_code, '' from twoexam " & vbCrLf & _
                      "where request_date = '" & Trim(GetText(vasID, intListRow, colReceDate)) & "' " & vbCrLf & _
                      "and exam_no = '" & Trim(GetText(vasID, intListRow, colReceno)) & "' and exam_code in (" & gAllExam & ") "
                res = db_select_Vas(gServer, SQL, vasTemp)
            End Select
            
            For x = 1 To vasTemp.DataRowCnt
                lsExamCode = Trim(GetText(vasTemp, x, 1))
                lsSubCode = Trim(GetText(vasTemp, x, 2))
                               
                SQL = "select examcode, subcode, equipcode, seqno, examname from equipexam " & vbCrLf & _
                      "where equipno = '" & gEquip & "' and examcode = '" & lsExamCode & "' and isnull(subcode, '') = '" & lsSubCode & "'"
                res = db_select_Col(gLocal, SQL)
                If res > 0 Then
                    lsEquipCode = Trim(gReadBuf(2))
                    lsSeqNo = Trim(gReadBuf(3))
                    lsExamName = Trim(gReadBuf(4))
                    
                    SQL = "select sampleno from pat_res " & vbCrLf & _
                          "where recedate = '" & Trim(GetText(vasList, i, 2)) & "' and pid = '" & Trim(GetText(vasList, i, 5)) & "' " & vbCrLf & _
                          "and receno = '" & Trim(GetText(vasList, i, 4)) & "' and examgubun = '" & chServer & "' and examcode = '" & lsExamCode & "' and subcode = '" & lsSubCode & "' "
                    res = db_select_Col(gLocal, SQL)
                    If res > 0 Then
                    Else
                        
                   gReadBuf(0) = ""
                   
                   SQL = "select result from pat_res " & vbCrLf & _
                         "where equipno = '" & gEquip & "' and sampleno = '" & Trim(GetText(vasID, intListRow, 3)) & "' " & vbCrLf & _
                         "and equipcode = '" & lsEquipCode & "' and examdate = '" & Format(dtpToday.Value, "yyyymmdd") & "' AND SENDFLAG = '0' "
                   res = db_select_Col(gLocal, SQL)
                   
                   result = Trim(gReadBuf(0))
                   
                   If res > 0 Then
                
                   SQL = "insert into pat_res(equipno, recedate, sampleno, receno, pid, " & vbCrLf & _
                         "pname, psex, page, pjumin, sendflag, " & vbCrLf & _
                         "examgubun,examdate, examcode, subcode, examname, equipcode, seqno, hospital,result) " & vbCrLf & _
                         "values('" & gEquip & "', '" & Trim(GetText(vasList, i, 2)) & "', '" & Trim(GetText(vasID, intListRow, 3)) & "' , " & vbCrLf & _
                         "'" & Trim(GetText(vasList, i, 4)) & "', '" & Trim(GetText(vasList, i, 5)) & "', '" & Trim(GetText(vasList, i, 6)) & "', " & vbCrLf & _
                         "'" & Trim(GetText(vasList, i, 7)) & "', '" & Trim(GetText(vasList, i, 8)) & "', '" & Trim(GetText(vasList, i, 9)) & "', " & vbCrLf & _
                         "'2', '" & chServer & "','" & Format(dtpToday, "YYYYMMDD") & "', '" & lsExamCode & "', " & vbCrLf & _
                         "'" & lsSubCode & "', '" & lsExamName & "', '" & lsEquipCode & "', '" & lsSeqNo & "', '미지정', '" & result & "')"
                   res = SendQuery(gLocal, SQL)
                        
                        End If
                    End If
                End If
            Next x
            vasID.RowHeight(intListRow) = 13
'           ' Make_Order intListRow
'
'            SQL = "update worklist set state = 1 where recedate = '" & Trim(GetText(vasList, i, 2)) & "' and receno = '" & Trim(GetText(vasList, i, 4)) & "' and equipno = '" & gEquip & "' "
'            res = SendQuery(gLocal, SQL)
            
        End If
        
    Next i
    
    For i = vasList.DataRowCnt To 1 Step -1
        vasList.Col = 1
        vasList.Row = i
        If vasList.Value = 1 Then
            DeleteRow vasList, i, i
        End If
        
    Next i
    
   
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
    dtpReceDate = Date
    dtpToday = Date
    dtpESDate = Date
    dtpEEDate = Date
    dtpEndDate = Date
    
End Sub

Private Sub cmdRSch_Click()
    Dim iRow As Long
    Dim sType As String

    ClearSpread vasRID
    ClearSpread vasRRes
    
    SQL = "SELECT '', recedate, sampleno, diskno, posno, receno, pid, pname, page, psex, pjumin, count(*), count(*), sendflag,EXAMGUBUN " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' AND EQUIPNO = '" & gEquip & "' and hospital = '" & cmbHosp.Text & "'"
    If optExamPart(0).Value = True Then
        SQL = SQL & "  AND SENDFLAG IN ('2', '3') "
    Else
        SQL = SQL & "  AND SENDFLAG IN ('0') "
    End If
    
    SQL = SQL & "GROUP BY recedate, sampleno, diskno, posno, receno, pid, pname, page, psex, pjumin, sendflag,EXAMGUBUN"
    res = db_select_Vas(gLocal, SQL, vasRID)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRID.RowHeight(-1) = 13
    
    For iRow = 1 To vasRID.DataRowCnt
        If Trim(GetText(vasRID, iRow, 14)) = "2" Then
            SetText vasRID, "결과", iRow, 14
            SetBackColor vasRID, iRow, iRow, 1, 1, 255, 250, 205
    
        ElseIf Trim(GetText(vasRID, iRow, 14)) = "3" Then
            SetBackColor vasRID, iRow, iRow, 1, 14, 202, 255, 112
            SetText vasRID, "완료", iRow, 14
        End If
    Next
    
    vasSort vasRID, 3
    vasRID.SetFocus
    vasRID.MaxRows = vasRID.DataRowCnt
    
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
                SetText vasRID, "실패", lRow, colState
            Else
                vasRID.Row = lRow
                vasRID.Col = 1
                vasRID.Value = 1
                
                SetBackColor vasRID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasRID, "완료", lRow, colState
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " SENDFLAG = '3' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND receno = '" & Trim(GetText(vasRID, lRow, colReceno)) & "' " & vbCrLf & _
                      " AND sampleno = '" & Trim(GetText(vasRID, lRow, colSampleNo)) & "' " & vbCrLf & _
                      " AND examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' "
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

Private Sub cmdWorklist_Click()
    Dim chServer
    Dim sType As String
    Dim iRow As Integer
    
    chServer = cmbPart.ListIndex
    ClearSpread vasList
    
    If chServer = dpGumjin1 Then
        SQL = "select '', a.request_date, '',a.exam_no, b.chart_no, b.person_name, '', '', b.personal_id, '0' " & vbCrLf & _
              "from totres a, total b " & vbCrLf & _
              "where a.request_date = b.request_date and a.exam_no = b.exam_no " & vbCrLf & _
              "  and a.request_date >= '" & Format(dtpReceDate, "yyyymmdd") & "' and a.request_date <= '" & Format(dtpEndDate, "yyyymmdd") & "' and a.result_value = '' " & vbCrLf & _
              "  and a.exam_code in (" & gAllExam & ") " & vbCrLf & _
              "group by a.request_date, a.exam_no, b.person_name, b.personal_id,b.chart_no " & vbCrLf & _
              "order by a.exam_no "
        res = db_select_Vas(gServer, SQL, vasList)
    ElseIf chServer = dpOCS Then
        SQL = "select '', a.년 + a.월 + a.일, '', a.챠트번호, a.오더일련번호, b.수진자명, '', '', b.주민등록번호, '1'  " & vbCrLf & _
              "from TB_진료검사 a, TB_인적사항 b " & vbCrLf & _
              "where a.챠트번호 = b.챠트번호 " & vbCrLf & _
              "  and a.년 + a.월 + a.일 >= '" & Format(dtpReceDate, "yyyymmdd") & "' and a.년 + a.월 + a.일 <= '" & Format(dtpEndDate, "yyyymmdd") & "' " & vbCrLf & _
              "  and a.검사코드 in (" & gAllExam_Ocs & " ) and a.오더일련번호 > '0' AND 상태 = '0' " & vbCrLf & _
              "  and a.검사종류 in (" & gAllExam1 & " ) " & vbCrLf & _
              "group by a.년 + a.월 + a.일,a.오더일련번호, a.챠트번호, b.수진자명, b.주민등록번호"
        res = db_select_Vas(gServer_OCS, SQL, vasList)

    ElseIf chServer = dpGumjin2 Then
        SQL = "select '', a.request_date, '',a.exam_no, b.chart_no, b.person_name, '', '', b.personal_id, '2'  " & vbCrLf & _
              "from twoexam a, total b, panjong2 c " & vbCrLf & _
              "where a.request_date = b.request_date and a.exam_no = b.exam_no " & vbCrLf & _
              "  and a.request_date = c.request_date and a.exam_no = c.exam_no " & vbCrLf & _
              "  and a.request_date >= '" & Format(dtpReceDate, "yyyymmdd") & "' and a.request_date <= '" & Format(dtpEndDate, "yyyymmdd") & "'" & vbCrLf & _
              "  and a.exam_code in (" & gAllExam & ") " & vbCrLf & _
              "group by a.request_date, a.exam_no, b.person_name, b.personal_id,b.chart_no " & vbCrLf & _
              "order by a.exam_no"
        res = db_select_Vas(gServer, SQL, vasList)
    Else
        SQL = "select '', a.request_date, '',a.exam_no, b.chart_no, b.person_name, '', '', b.personal_id, '0' " & vbCrLf & _
              "from totres a, total b " & vbCrLf & _
              "where a.request_date = b.request_date and a.exam_no = b.exam_no " & vbCrLf & _
              "  and a.request_date >= '" & Format(dtpReceDate, "yyyymmdd") & "' and a.request_date <= '" & Format(dtpEndDate, "yyyymmdd") & "' and a.result_value = '' " & vbCrLf & _
              "  and a.exam_code in (" & gAllExam & ") " & vbCrLf & _
              "group by a.request_date, a.exam_no, b.person_name, b.personal_id,b.chart_no " & vbCrLf & _
              "order by a.exam_no "
        res = db_select_Vas(gServer, SQL, vasList)
        
        'SQL = "select '', a.년 + a.월 + a.일, '', a.챠트번호, a.오더일련번호, b.수진자명, '', '', b.주민등록번호, '1'  " & vbCrLf & _
              "from TB_진료검사 a, TB_인적사항 b " & vbCrLf & _
              "where a.챠트번호 = b.챠트번호 " & vbCrLf & _
              "  and a.년 + a.월 + a.일 >= '" & Format(dtpReceDate, "yyyymmdd") & "' and a.년 + a.월 + a.일 <= '" & Format(dtpEndDate, "yyyymmdd") & "' " & vbCrLf & _
              "  and a.검사코드 in (" & gAllExam_Ocs & " ) and a.오더일련번호 > '0' AND 상태 = '0' " & vbCrLf & _
              "  and a.검사종류 in (" & gAllExam & " ) " & vbCrLf & _
              "group by a.년 + a.월 + a.일,a.오더일련번호, a.챠트번호, b.수진자명, b.주민등록번호"
        'res = db_select_Vas(gServer_OCS, SQL, vasList, vasList.DataRowCnt + 1)
        
        'SQL = "select '', a.request_date, '',a.exam_no, b.chart_no, b.person_name, '', '', b.personal_id, '2'  " & vbCrLf & _
              "from twoexam a, total b, panjong2 c " & vbCrLf & _
              "where a.request_date = b.request_date and a.exam_no = b.exam_no " & vbCrLf & _
              "  and a.request_date = c.request_date and a.exam_no = c.exam_no " & vbCrLf & _
              "  and a.request_date >= '" & Format(dtpReceDate, "yyyymmdd") & "' and a.request_date <= '" & Format(dtpEndDate, "yyyymmdd") & "'" & vbCrLf & _
              "  and a.exam_code in (" & gAllExam & ") " & vbCrLf & _
              "group by a.request_date, a.exam_no, b.person_name, b.personal_id,b.chart_no " & vbCrLf & _
              "order by a.exam_no"
        'res = db_select_Vas(gServer, SQL, vasList, vasList.DataRowCnt + 1)
    End If
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasList.MaxRows = vasList.DataRowCnt
    
'    For iRow = 1 To vasListTemp.DataRowCnt
'
'        SQL = "select receno from worklist where recedate = '" & Trim(GetText(vasListTemp, iRow, 2)) & "' and receno = '" & Trim(GetText(vasListTemp, iRow, 4)) & "' and equipno = '" & gEquip & "' "
'        res = db_select_Col(gLocal, SQL)
'        If res > 0 Then
'        Else
'            CalAgeSex Trim(GetText(vasList, iRow, 9)), Trim(Format(dtpToday.Value, "yyyy-mm-dd"))
'            SQL = "insert into worklist(recedate, receno, pid, pname, page, psex, pjumin, state, examgubun,equipno) " & vbCrLf & _
'                  "values('" & Trim(GetText(vasListTemp, iRow, 2)) & "','" & Trim(GetText(vasListTemp, iRow, 4)) & "'," & vbCrLf & _
'                  "'" & Trim(GetText(vasListTemp, iRow, 5)) & "','" & Trim(GetText(vasListTemp, iRow, 6)) & "'," & vbCrLf & _
'                  "'" & gPatGen.Age & "','" & gPatGen.Sex & "'," & vbCrLf & _
'                  "'" & Trim(GetText(vasListTemp, iRow, 9)) & "',0,'" & Trim(GetText(vasListTemp, iRow, 10)) & "', '" & gEquip & "')"
'            res = SendQuery(gLocal, SQL)
'        End If
'
'
''        vasList.Row = iRow
''        vasList.Col = 1
''        vasList.Value = 1
'    Next iRow
'
'    SQL = "select '', recedate, '', receno, pid, pname, psex, page, pjumin, examgubun from worklist" & vbCrLf & _
'          "where recedate >= '" & Format(dtpReceDate, "yyyymmdd") & "' and recedate <= '" & Format(dtpEndDate, "yyyymmdd") & "' " & vbCrLf & _
'          "and equipno = '" & gEquip & "' "
'    SQL = SQL & " and state = 0 "
'
'    If cmbPart.ListIndex = 3 Then
'    Else
'        SQL = SQL & " and examgubun = '" & chServer & "'"
'    End If
    
'    res = db_select_Vas(gLocal, SQL, vasList)
    
    For iRow = 1 To vasList.DataRowCnt
        CalAgeSex Trim(GetText(vasList, iRow, 9)), Trim(Format(dtpToday.Value, "yyyy-mm-dd"))
        SetText vasList, gPatGen.Sex, iRow, 7
        SetText vasList, gPatGen.Age, iRow, 8
        SetText vasList, chServer, iRow, 10
        
'        vasList.Row = iRow
'        vasList.Col = 1
'        vasList.Value = 1
    Next iRow
    vasList.RowHeight(-1) = 13

End Sub

Private Sub Command1_Click()

'    SQL = "SELECT '', recedate, sampleno, pid, pname " & vbCrLf & _
'          "FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(DTPicker2, "YYYYMMDD") & "' AND EQUIPNO = '" & gEquip & "' "
'    SQL = SQL & "GROUP BY recedate, sampleno,  receno, pid, pname"
'    res = db_select_Vas(gLocal, SQL, vasRID)
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
End Sub

Private Sub Command16_Click()



Dim intIdx      As Integer
Dim strSrcfile  As String
Dim varBuffer   As Variant
Dim strBuffer   As String

    FileURIT.Path = gMachPath
    FileURIT.Refresh
    
    For intIdx = 0 To FileURIT.ListCount - 1
        
        FileURIT.ListIndex = intIdx
        
        '===== 조회기간에 맞는것만 1 =================================================================
        If FileURIT.FileName = "NameResult.txt" Then
            strSrcfile = FileURIT.Path & "\" & FileURIT.FileName   ' 원본 파일 이름을 정의합니다.
            
            Open strSrcfile For Input As #9
        
            strBuffer = ""
        
            Do While Not EOF(9)
                strBuffer = strBuffer & Input(1, #9)
                'Line Input #9, sLine
                'strBuffer = strBuffer & sLine & vbCr
            Loop
        
            Close #9
            
            varBuffer = Split(strBuffer, vbCr)
            
            For i = 0 To UBound(varBuffer) - 1
                Debug.Print varBuffer(i)
            Next
        End If
    Next

    u411 Mid(txtTest, 2)
    txtTest = ""


End Sub

Sub SendOrder()
    Dim sSendOrder As String
    
    If Len(gOrderMessage) > 240 Then
        
        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If
        
        sSendOrder = CStr(gOrderCnt) & Left(gOrderMessage, 240) & chrETB
        gOrderMessage = Mid(gOrderMessage, 241)
        
        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
        SaveQuery sSendOrder, 1
        
        gOrderCnt = gOrderCnt + 1
        comSend = "stENQ"
        
        gPreMsg = sSendOrder
        
        Save_Raw_Data "[Tx]" & gPreMsg
        MSComm1.Output = sSendOrder
        
    Else
        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If
        
        sSendOrder = CStr(gOrderCnt) & gOrderMessage & chrETX
        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
                
        gOrderMessage = ""
        comSend = "stOrder"
        
        gPreMsg = sSendOrder
        
        Save_Raw_Data "[Tx]" & gPreMsg
        MSComm1.Output = sSendOrder
    End If
End Sub

Private Sub Command3_Click()
    Dim iRow As Long
    Dim sType As String
    Dim chServer
    
    chServer = cmbServerType.ListIndex
    ClearSpread vasID
    ClearSpread vasRes
    
    SQL = "SELECT '', recedate, sampleno, diskno, posno, receno, pid, pname, page, psex, pjumin, count(*), count(*), sendflag,EXAMGUBUN " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' AND EQUIPNO = '" & gEquip & "' "
    SQL = SQL & "GROUP BY recedate, sampleno, diskno, posno, receno, pid, pname, page, psex, pjumin, sendflag,EXAMGUBUN"
    res = db_select_Vas(gLocal, SQL, vasID)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
End Sub

Private Sub Command4_Click()
    Dim i As Integer
    Dim Response, Help
    
    Response = MsgBox("삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!", Help, 100)
    
    If Response = vbYes Then
        For i = vasRID.DataRowCnt To 1 Step -1
            vasRID.Col = 1
            vasRID.Row = i
            If vasRID.Value = 1 Then
            
                SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                      "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
                      "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      "  AND receno = '" & Trim(GetText(vasRID, vasRID.Row, colReceno)) & "' " & vbCrLf & _
                      "  and sampleno = '" & Trim(GetText(vasRID, vasRID.Row, colSampleNo)) & "' "
                res = SendQuery(gLocal, SQL)
        
                DeleteRow vasRID, i, i
            End If
        Next i
    End If
End Sub

Private Sub dtpExamDate_Change()
    Dim lsExamDate As String
    lsExamDate = Format(dtpExamDate, "yyyymmdd")
    
    HospChk Format(dtpExamDate, "yyyymmdd")
End Sub


Private Sub HospChk(asDate As String)
    Dim lsExamDate As String
    
    lsExamDate = asDate
    
    cmbHosp.Clear
    lsExamDate = Format(dtpExamDate, "yyyymmdd")
    SQL = "select distinct hospital from pat_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' and examdate = '" & lsExamDate & "' "
    res = db_select_Combo(gLocal, SQL, cmbHosp)
    cmbHosp.Text = "미지정"
End Sub


Private Sub Form_Load()
    Dim sDate As String
    
'''    If ProcessCHK("IF_U411.exe") = True Then
'''        MsgBox "인터페이스 프로그램이 실행중입니다.", vbOKOnly, "경고"
'''        Unload Me
'''    End If
    
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    Me.Height = 11520
    Me.Width = 15435
    
    cmdIFClear_Click
    cmdRClear_Click
    
    GetSetup
    
'    MSComm1.CommPort = gSetup.gPort
'    MSComm1.RTSEnable = gSetup.gRTSEnable
'    MSComm1.DTREnable = gSetup.gDTREnable
'    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
'
'    If MSComm1.PortOpen = False Then
'        MSComm1.PortOpen = True
'    End If
    
    '진료
    If Not Connect_Server_OCS Then
        MsgBox "연결되지 않았습니다."
        cn_ServerOCS_Flag = False
        Exit Sub
    Else
        cn_ServerOCS_Flag = True
    End If
    
    '검진
    If Not Connect_Server Then
        MsgBox "연결되지 않았습니다."
        cn_Server_Flag = False
        Exit Sub
    Else
        cn_Server_Flag = True
    End If
    
    '로컬
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If

    GetExamCode
    dtpToday = Date
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -365), "yyyymmdd")
    
    SQL = "delete from pat_res where examdate < '" & sDate & "'"
    res = SendQuery(gLocal, SQL)
    
    stInterface.Tab = 0
    
    cmbPart.AddItem "AlLISS", 0
    cmbPart.AddItem "진료", 1
    cmbPart.AddItem "2차검진", 2
    cmbPart.AddItem "[전체]", 3
    
    
    cmbServerType.AddItem "AlLISS", 0
    cmbServerType.AddItem "진료", 1
    cmbServerType.AddItem "2차검진", 2
    cmbServerType.AddItem "[전체]", 3
    
    cmbPart.ListIndex = 3
    cmbServerType.ListIndex = 3
    
    HospChk Format(dtpExamDate, "yyyymmdd")
    
'    Dim i As Integer
'
'    SQL = "select equipcode, examname, seqno from equipexam where equipno = '" & gEquip & "' and examgubun = '0'"
'    res = db_select_Vas(gLocal, SQL, vasID)
'
'    For i = 1 To vasID.DataRowCnt
'        SQL = "update equipexam set examname = '" & Trim(GetText(vasID, i, 2)) & "', seqno = '" & Trim(GetText(vasID, i, 3)) & "' " & vbCrLf & _
'              "where equipno = '" & gEquip & "' and equipcode = '" & Trim(GetText(vasID, i, 1)) & "'"
'        res = SendQuery(gLocal, SQL)
'    Next

    FileURIT.Path = gMachPath
    
End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
Dim lsBarcode   As String
Dim lsPID       As String
Dim lsReceNo    As String
Dim sRes        As String

    Get_Sample_Info = -1
    
    '샘플 환자 정보 가져오기
    
    lsBarcode = Trim(GetText(vasID, asRow, colReceno))   '샘플 바코드 번호
    
'    SQL = "SELECT A.SPCM_NO, B.PID, B.PT_NM, B.FRRN, B.SEX_CD " & vbCrLf & _
'          "FROM SY_MSLDEXMNRSLT A, SY_PCPMPT B " & vbCrLf & _
'          "Where A.PID = B.PID " & vbCrLf & _
'          "AND A.SPCM_NO = '" & lsBarcode & "' "
    SQL = "select '', pid, pname, pjumin, psex from pat_res where recedate = '" & Trim(GetText(vasID, asRow, colReceDate)) & "' and receno = '" & Trim(GetText(vasID, asRow, colReceno)) & "'"
    res = db_select_Col(gLocal, SQL)
    
    If res > 0 Then
        SetText vasID, Trim(gReadBuf(1)), asRow, colPID
        SetText vasID, Trim(gReadBuf(2)), asRow, colPName
        SetText vasID, Trim(gReadBuf(3)), asRow, colPJumin
        SetText vasID, Trim(gReadBuf(4)), asRow, colPSex
    End If

    Get_Sample_Info = 1

End Function

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

'Function GetExamCode() As Integer
'    Dim i, j As Long
'
'    ClearSpread vasTemp
'    GetExamCode = -1
'    gAllExam = ""
'    SQL = "Select equipcode, examcode, examname, range, seqno " & vbCrLf & _
'          "From equipexam " & vbCrLf & _
'          "Where equipno = '" & gEquip & "' " & vbCrLf & _
'          "and examgubun = '1' " & vbCrLf & _
'          "order by seqno "
'    res = db_select_Vas(gLocal, SQL, vasTemp)
'    If res > 0 Then
'        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 6)
'    ElseIf res < 0 Then
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    For i = 1 To vasTemp.DataRowCnt
'        If i = 1 Then
'            gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
'        Else
'            gAllExam = gAllExam & ",'" & Trim(GetText(vasTemp, i, 2)) & "'"
'        End If
'
'        gArrEquip(i, 1) = i
'        For j = 1 To 5
'            gArrEquip(i, j + 1) = Trim(GetText(vasTemp, i, j))
'        Next j
'    Next i
'
'    ClearSpread vasTemp
'    gAllExam_Ocs = ""
'    SQL = "Select equipcode, examcode, subcode, examname, range, seqno " & vbCrLf & _
'          "From equipexam " & vbCrLf & _
'          "Where equipno = '" & gEquip & "' " & vbCrLf & _
'          "and examgubun = '2' " & vbCrLf & _
'          "order by seqno "
'    res = db_select_Vas(gLocal, SQL, vasTemp)
'    If res > 0 Then
'        ReDim gArrEquip_Ocs(1 To vasTemp.DataRowCnt, 1 To 6)
'    Else
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    For i = 1 To vasTemp.DataRowCnt
'        If i = 1 Then
'            gAllExam_Ocs = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
'        Else
'            gAllExam_Ocs = gAllExam_Ocs & ",'" & Trim(GetText(vasTemp, i, 2)) & "'"
'        End If
'
'        gArrEquip_Ocs(i, 1) = i
'        For j = 1 To 5
'            gArrEquip_Ocs(i, j + 1) = Trim(GetText(vasTemp, i, j))
'        Next j
'    Next i
'
'    For i = 1 To vasTemp.DataRowCnt
'        If i = 1 Then
'            gAllExam1 = "'" & Trim(GetText(vasTemp, i, 3)) & "'"
'        Else
'            gAllExam1 = gAllExam1 & ",'" & Trim(GetText(vasTemp, i, 3)) & "'"
'        End If
'    Next i
'
'    GetExamCode = 1
'End Function

Function GetExamCode() As Integer
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
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
    frmOrderCode.Show 1
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show 1
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

Private Sub MnvasIdDel_Click()
    Dim i As Integer
    
    For i = vasID.DataRowCnt To 1 Step -1
        vasID.Col = 1
        vasID.Row = i
        If vasID.Value = 1 Then
            DeleteRow vasID, i, i
        End If
        
    Next i
    
End Sub

Private Sub MnvasListDel_Click()
    Dim i As Integer
    
    For i = vasList.DataRowCnt To 1 Step -1
        vasList.Col = 1
        vasList.Row = i
        If vasList.Value = 1 Then
            DeleteRow vasList, i, i
        End If
        
    Next i
    
End Sub

Private Sub MnvasRIDDel_Click()
    Dim i As Integer
    
    For i = vasRID.DataRowCnt To 1 Step -1
        vasRID.Col = 1
        vasRID.Row = i
        If vasRID.Value = 1 Then
            DeleteRow vasRID, i, i
        End If
        
    Next i
End Sub

Private Sub MSComm1_OnComm()
        Dim lsChar As String
    
    lsChar = MSComm1.Input
        
    Select Case lsChar
    Case chrENQ
    
        Save_Raw_Data "[RX]" & lsChar
        
        MSComm1.Output = chrACK
        Save_Raw_Data "[TX]" & chrACK
    Case chrSTX
        txtData.Text = ""
        txtData.Text = chrSTX
        
    Case chrLF
        txtData.Text = txtData.Text & lsChar
        Save_Raw_Data "[RX]" & txtData.Text
        
        u411 Mid(txtData, 2)
        
        MSComm1.Output = chrACK
        Save_Raw_Data "[TX]" & chrACK
        
        txtData.Text = ""
    Case chrEOT
        Save_Raw_Data "[RX]" & lsChar
        
    Case Else
        txtData.Text = txtData.Text & lsChar
    End Select
End Sub

Sub u411(asData As String)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Dim ResultTbl(1 To 50) As String        'Array에 담기
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim sCnt As String
    
    Dim sDate As String
    Dim sRefFlag As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sPanicFlag As String
    Dim sDeltaFlag As String
    Dim sReceNo As String
    Dim sReceDate  As String
    Dim sPID As String
    Dim sPName As String
    Dim sJumin As String
    Dim sPSex As String
    Dim sPage As String
    Dim sTestID As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sResult1 As String
    Dim sExamDate As String
    Dim sExamName As String
    Dim sCount As Integer
    Dim sSeqNo As String
    Dim liRet  As Integer
    Dim sSpace As Integer
    
    Dim sSampleType As String
    
    Dim sLevelNo As String
    
    Dim lsTemp1 As String
    Dim jRow As Integer
    
    If asData = "" Then
        Exit Sub
    End If
    
    TablePtr = 1
' ----- for start
    For j = 1 To Len(asData)
        If (Mid(asData, j, 1) = "|") Then
            TablePtr = TablePtr + 1
            ResultTbl(TablePtr) = " "
        Else
            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
        End If
    Next j
' ------- for end
    
    If Mid(ResultTbl(1), 2, 1) = "H" Then           'Header Record
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "C" Then           'Comment Record
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "L" Then
        If MnTransAuto.Checked = True Then
            If Insert_Data(llRow) = 1 Then
                SQL = "update pat_res set sendflag = '3'" & vbCrLf & _
                      "WHERE examdate = '" & Format(CDate(dtpToday.Value), "yyyymmdd") & "' " & vbCrLf & _
                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND receno = '" & Trim(GetText(vasID, llRow, colReceno)) & "' " & vbCrLf & _
                      "  AND PID = '" & Trim(GetText(vasID, llRow, colPID)) & "' " & vbCrLf & _
                      "  AND SAMPLENO = '" & Trim(GetText(vasID, llRow, colSampleNo)) & "' " & vbCrLf & _
                res = SendQuery(gLocal, SQL)

                SetBackColor vasID, llRow, llRow, colCheckBox, colCheckBox, 202, 255, 112
                SetText vasID, "완료", llRow, colState
            Else
                SetBackColor vasID, llRow, llRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasID, "실패", llRow, colState
            End If
        End If
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "Q" Then           'Query Record
    End If
    
    
    If (Mid(ResultTbl(1), 2, 1) = "O") Then         'Test Order Record
        sSeqNo = Trim(ResultTbl(4))
        i = InStr(1, sSeqNo, "^")
        sSeqNo = Mid(sSeqNo, 1, i - 1)
        
        llRow = -1
        For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colSampleNo)) = sSeqNo Then
                llRow = i
                Exit For
            End If
        Next i
        
        If llRow = -1 Then
            For i = 1 To vasID.DataRowCnt
            If Trim(GetText(vasID, i, colSampleNo)) = "" Then
                llRow = i
                Exit For
            End If
        Next i
        End If
        
        
        If llRow = -1 Then
            llRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < llRow Then
                vasID.MaxRows = llRow
            End If
        End If
        
        
        SetText vasID, sSeqNo, llRow, colSampleNo
                            
        SetText vasID, "수신중", llRow, colState
        ClearSpread vasRes
        
    End If

    If (Mid(ResultTbl(1), 2, 1) = "R") Then
        sTmp = ResultTbl(3)
        i = InStr(1, Trim(sTmp), "^")
       
        If i > 0 Then
            sTestID = Mid(Trim(sTmp), i + 1)
        Else
            sTestID = ""
        End If
        
        sTmp = ResultTbl(4)
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            sResult = Trim(Mid(sTmp, 1, i - 1))
        Else
            sResult = Trim(sTmp)
        End If
        
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            sResult1 = Trim(Mid(sTmp, i + 1))
        Else
            sResult1 = Trim(sTmp)
        End If
        
        If sResult1 = "+" Then
            sResult1 = "양성(+)"
        ElseIf sResult1 = "++" Then
            sResult1 = "양성(++)"
        ElseIf sResult1 = "+++" Then
            sResult1 = "양성(+++)"
        ElseIf sResult1 = "++++" Then
            sResult1 = "양성(++++)"
        ElseIf sResult1 = "+++++" Then
            sResult1 = "양성(+++++)"
        ElseIf sResult1 = "TR" Then
            sResult1 = "trace"
        ElseIf sResult1 = "TRACE" Then
            sResult1 = "trace"
        ElseIf sResult1 = "tr" Then
            sResult1 = "trace"
        ElseIf sResult1 = "+-" Then
            sResult1 = "약양성(+-)"
        End If
        
        If sResult = "" Then
            Exit Sub
        End If
        If sTestID = "SG" Then
            sResult = sResult
        ElseIf sTestID = "pH" Then
            sResult = Format(sResult, "0.0#")
        ElseIf sTestID = "UBG" Then
            If sResult = "norm" Then
                sResult = "음성"
            Else
                sResult = sResult1
                
                If sResult1 = "neg" Then
                    sResult = "음성"
                ElseIf sResult1 = "norm" Then
                    sResult = "음성"
                ElseIf sResult1 = "pos" Then
                    sResult = "양성"
                Else
                    sResult = sResult1
                End If
            End If
        Else
            If sResult1 = "neg" Then
                sResult = "음성"
            ElseIf sResult1 = "pos" Then
                sResult = "양성"
            Else
                sResult = sResult1
            End If
        End If
        
        gReadBuf(0) = ""
        gReadBuf(1) = ""
        SQL = "Select count(equipcode) From EquipExam" & vbCrLf & _
              " Where Equipno = '" & gEquip & "' "
        res = db_select_Col(gLocal, SQL)
        vasRes.MaxRows = gReadBuf(0)
'        vasRes.MaxRows = 1
        ClearSpread vasTemp
        
        If vasRes.DataRowCnt = 0 Then
            vasRes.MaxRows = 1
        Else
            vasRes.MaxRows = vasRes.DataRowCnt + 1
        End If

        j = vasRes.DataRowCnt + 1
        
        If sResult <> "" Then
            
            gReadBuf(0) = ""
            gReadBuf(1) = ""
            gReadBuf(2) = ""
        
            SQL = "select examcode, examname, subcode,seqno from PAT_RES" & vbCrLf & _
                  "Where Equipno = '" & gEquip & "' And EQUIPCODE = '" & Trim(sTestID) & "' " & vbCrLf & _
                  "AND PID = '" & Trim(GetText(vasID, llRow, colPID)) & "'  " & vbCrLf & _
                  "AND examdate = '" & Format(dtpToday.Value, "YYYYMMDD") & "' " & vbCrLf & _
                  "AND RECENO = '" & Trim(GetText(vasID, llRow, colReceno)) & "' "
            res = db_select_Col(gLocal, SQL)
            If res > 0 Then
            SetText vasRes, Trim(sTestID), j, colEquipCode      '장비코드
            SetText vasRes, gReadBuf(0), j, colExamCode         '검사코드
            SetText vasRes, gReadBuf(1), j, colExamName         '검사명
            SetText vasRes, gReadBuf(2), j, colSubCode       '검사명
            SetText vasRes, gReadBuf(3), j, colSeq
            SetText vasRes, sResult, j, colResult               '검사결과
            
            If Trim(GetText(vasID, llRow, colReceno)) <> "" Then
               Save_Local_One llRow, j, "2"
            Else
               Save_Local_One llRow, j, "0"
            End If
            
            Else
            SQL = "select examcode, examname, subcode,seqno from equipexam" & vbCrLf & _
                  "Where Equipno = '" & gEquip & "' And EQUIPCODE = '" & Trim(sTestID) & "' " & vbCrLf & _
                  "AND examgubun = '2'  "
            res = db_select_Col(gLocal, SQL)
            SetText vasRes, Trim(sTestID), j, colEquipCode      '장비코드
            SetText vasRes, gReadBuf(0), j, colExamCode         '검사코드
            SetText vasRes, gReadBuf(1), j, colExamName         '검사명
            SetText vasRes, gReadBuf(2), j, colSubCode       '검사명
            SetText vasRes, gReadBuf(3), j, colSeq
            SetText vasRes, sResult, j, colResult               '검사결과
            
            Save_Local_One llRow, j, "0"
            
           End If
        End If

            
        SetText vasID, "결과", llRow, colState
        SetBackColor vasID, llRow, llRow, 1, 1, 255, 250, 205
        '==============================================================================================
    End If

End Sub

Function Proc_Result(asOrd As String, ByVal argSpread As vaSpread) As Integer
    Dim i, j, k, iArr, lResRow As Long
    Dim iStr As Integer
    Dim iCnt As Integer
    
    Dim sGubun As String
    Dim lsData As String
    Dim lsTemp As String
    Dim lsSampleType As String
    Dim lsSpecimenID As String
    Dim lsOrder As String
    Dim lsID As String
    Dim lsRackID As String
    Dim lsPosNO As String
    Dim lsPriority As String
    
    Dim lsExamCode As String
    Dim lsExamDate As String
    Dim lsEquipCode As String
    Dim lsResult As String
    Dim lsUnit As String
    Dim lsRef As String
    Dim lsState As String
    Dim lsComment As String
    Dim lsPoint As String
    Dim sBarcode As String
    Dim lsRefer As String
    Dim iRow As Integer
    
    Dim sCnt As String
    
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsGubun As String
    
    Dim sExamCodeAll As String
    Dim sTmpStr As String
    
    Dim lsReqDate As String
    Dim lsReceNo As String
    Dim lsExamGubun As Boolean
    Dim lsResChk As String
    
    
    lsResChk = "2"
    
    Proc_Result = -1
    lsExamGubun = False
    
'    gOrd.OrderCnt = 0
'    gOrd.OrderText = ""
    lsData = asOrd
    i = 0
    iStr = 1
    iCnt = 0
    lsID = ""
    i = InStr(iStr, lsData, "|")
    Do While i > 0
        iCnt = iCnt + 1
        
        lsTemp = Mid(lsData, iStr, i - iStr)
        lsData = Mid(lsData, i + 1)
        
        Select Case iCnt
        Case 3
            sBarcode = Trim(lsTemp)
        Case 4
            lsSpecimenID = lsTemp
        Case 5
            lsOrder = lsTemp
        Case 6
            lsPriority = lsTemp
            Exit Do
        Case 23
            lsExamDate = lsTemp
            Exit Do
        Case Else
        End Select
        
        lsTemp = ""
        i = InStr(iStr, lsData, "|")
    Loop
    
    'lsExamDate = Left(lsExamDate, 4) & "-" & Mid(lsExamDate, 5, 2) & "-" & Mid(lsExamDate, 7, 2) & " " & Mid(lsExamDate, 9, 2) & ":" & Mid(lsExamDate, 11, 2) & ":" & Mid(lsExamDate, 13, 2)
    lsExamDate = Format(CDate(GetDateFull), "yyyy-mm-dd hh:nn:ss")
    
    i = InStr(1, lsSpecimenID, "^")
    If i > 0 Then
        lsID = Trim(Left(lsSpecimenID, i - 1))
        lsSpecimenID = Mid(lsSpecimenID, i + 1)
        'lsID = Trim(Left(lsSpecimenID, 13))
        'lsSpecimenID = Mid(lsSpecimenID, 14)
        i = InStr(1, lsSpecimenID, "^")
        If i > 0 Then
            lsRackID = Left(lsSpecimenID, i - 1)
'            lsPosNO = Trim(Mid(lsSpecimenID, i + 1))
'            lsID = Trim(Left(lsSpecimenID, i - 1))
            lsSpecimenID = Mid(lsSpecimenID, i + 1)
            i = InStr(1, lsSpecimenID, "^")
            If i > 0 Then
                lsPosNO = Left(lsSpecimenID, i - 1)
'                lsSampleType = Trim(Left(lsSpecimenID, i - 1))
                lsSampleType = Mid(lsSpecimenID, i + 1)
                'lsRackID = Mid(lsSpecimenID, 1, i - 1)
                i = InStr(1, lsSpecimenID, "^")
'                If i > 0 Then
'                    lsRackID = Left(lsSpecimenID, i - 1)
'                    lsPosNO = Trim(Mid(lsSpecimenID, i + 1))
'                End If
            End If
        End If
    End If
    lsID = Trim(Mid(lsID, 1, 15))
    i = InStr(1, sBarcode, "-")
    lsReqDate = ""
    lsReceNo = ""
    If i > 0 Then
        lsReqDate = "20" & Mid(sBarcode, 1, i - 1)
        lsReceNo = Mid(sBarcode, i + 1)
        lsExamGubun = True
    End If
    
        
'''    If Left(lsRackID, 1) = "4" Then
'''        iRow = -1
'''        For i = 1 To vasID.DataRowCnt
'''            If Trim(GetText(vasID, i, colRack)) = lsRackID And Trim(GetText(vasID, i, colPos)) = lsPosNO Then
'''                iRow = i
'''                Exit For
'''            End If
'''        Next i
'''        If iRow = -1 Then
'''            iRow = vasID.DataRowCnt + 1
'''            If iRow > vasID.MaxRows Then
'''                vasID.MaxRows = iRow + 1
'''            End If
'''
'''            SetText vasID, lsID, iRow, colBarCode
'''        End If
'''
'''    Else
        If lsExamGubun = True Then
        
            iRow = -1
            For i = vasID.DataRowCnt To 1 Step -1
                If Trim(GetText(vasID, i, colReceDate)) = lsReqDate And Trim(GetText(vasID, i, colReceno)) = lsReceNo Then
                    iRow = i
                    Exit For
                End If
            Next i
            
    '        If iRow = -1 Then
    '            For i = 1 To vasID.DataRowCnt
    '                If Trim(GetText(vasID, i, colRack)) = lsRackID And Trim(GetText(vasID, i, ColPos)) = lsPosNO Then
    '                    iRow = i
    '                    Exit For
    '                End If
    '            Next i
    '        End If
            
            If iRow = -1 Then
                iRow = vasID.DataRowCnt + 1
                If iRow > vasID.MaxRows Then
                    vasID.MaxRows = iRow + 1
                End If
                
                SetText vasID, lsID, iRow, colSampleNo
                SetText vasID, lsReqDate, iRow, colReceDate
                SetText vasID, lsReceNo, iRow, colReceno
                
            End If
        Else
            iRow = -1
            
            For i = 1 To vasID.DataRowCnt
                If CLng(Trim(GetText(vasID, i, colSampleNo))) = lsID Then
                    iRow = i
                    Exit For
                End If
            Next i
            If iRow = -1 Then
                iRow = vasID.DataRowCnt + 1
                If iRow > vasID.MaxRows Then
                    vasID.MaxRows = iRow + 1
                End If
                
                SetText vasID, lsID, iRow, colSampleNo
                SetText vasID, lsReqDate, iRow, colReceDate
                SetText vasID, lsReceNo, iRow, colReceno
                
            End If
        End If
'''    End If
         
    ClearSpread vasRes, 1, 1
    sExamCodeAll = ""
    
    vasID.SetText colRack, iRow, lsRackID
    vasID.SetText colPos, iRow, lsPosNO
    
    
'''    If cn_Server_Flag = True Then
'''        DisConnect_Server
'''    Else
'''        Connect_Server
'''    End If
    
    If Trim(GetText(vasID, iRow, colPName)) = "" And lsExamGubun = True Then
        Get_Sample_Info iRow
    End If
    
'            If Trim(GetText(vasID, iRow, colPName)) = "" Then
'                Get_Sample_Info iRow
'            End If
    
    ClearSpread vasTemp
    If lsExamGubun = False Then
    SQL = "select distinct examcode, equipcode from equipexam where equipno = '" & gEquip & "'"
    res = db_select_Vas(gLocal, SQL, vasTemp)
    Else
''    SQL = " select exam_code from totres " & vbCrLf & _
''          "where request_date = '" & Trim(GetText(vasID, iRow, colReqDate)) & "' and exam_no = '" & CLng(Trim(GetText(vasID, iRow, colSampleNo))) & "' " & vbCrLf & _
''          "and exam_code in (" & gAllExam & ") "
    SQL = "select distinct examcode, equipcode from pat_res where equipno = '" & gEquip & "' and recedate = '" & lsReqDate & "' and receno = '" & lsReceNo & "'"
    
    res = db_select_Vas(gLocal, SQL, vasTemp)
    End If
    
            
    
    For i = 1 To vasTemp.DataRowCnt
        If sExamCodeAll = "" Then
            sExamCodeAll = "'" & Trim(GetText(vasTemp, i, 1)) & "'"
        Else
            sExamCodeAll = sExamCodeAll & ", '" & Trim(GetText(vasTemp, i, 1)) & "'"
        End If
    Next i
    
    For iArr = 1 To argSpread.DataRowCnt
        iStr = 1
        iCnt = 0
        lsData = GetText(argSpread, iArr, 1)
        If lsData <> "" Then
            i = InStr(iStr, lsData, "|")
            Do While i > 0
                iCnt = iCnt + 1
                lsTemp = Mid(lsData, iStr, i - iStr)
                lsData = Mid(lsData, i + 1)
                
                Select Case iCnt
                Case 3
                    lsEquipCode = lsTemp
                    j = InStr(1, lsEquipCode, "/")
                    If j > 0 Then
                        lsEquipCode = Mid(lsEquipCode, 4, j - 4)
                    Else
                    lsEquipCode = Mid(lsEquipCode, 4, Len(lsEquipCode) - 4)
                    End If
                Case 4
                    lsResult = lsTemp
                    If InStr(1, lsResult, "^") > 0 Then
                        lsResult = Mid(lsResult, InStr(1, lsResult, "^") + 1)
                    End If
                    'Exit Do
                Case 5
                    lsUnit = lsTemp
                Case 7
                    lsRef = lsTemp
                    If UCase(lsRef) = "N" Then lsRef = ""
                    Exit Do
                Case 9
                    lsState = lsTemp
                    Exit Do
                Case Else
                End Select
                
                lsTemp = ""
                i = InStr(iStr, lsData, "|")
            
            Loop
            
            lResRow = vasRes.DataRowCnt + 1
            If vasRes.MaxRows < lResRow Then
                vasRes.MaxRows = lResRow
            End If
            
                        
            lsExamCode = ""
            
            If vasRes.MaxRows < lResRow Then vasRes.MaxRows = lResRow
            
            
'            vasRes.SetText 2, lResRow, lsID
            
            
                
            
            SQL = "Select examcode, examname, range, reflow, refhigh, resgubun  From equipexam" & vbCrLf & _
                  " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                  "  And equipcode = '" & lsEquipCode & "'" & vbCrLf & _
                  "  and examcode in (" & sExamCodeAll & ") "
            
            res = db_select_Col(gLocal, SQL)
            If (res = 1) And (gReadBuf(0) <> "") Then
                If IsNumeric(lsResult) Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                    lsPoint = Trim(gReadBuf(2))
                    lsResChk = Trim(gReadBuf(5))
                    
'                    If Trim(GetText(vasID, iRow, colPSex)) = "M" Then
'                        SQL = "select refer_low_man, refer_high_man from totres " & vbCrLf & _
'                              "where request_date = '" & Trim(GetText(vasID, iRow, colReqDate)) & "' " & vbCrLf & _
'                              "and exam_no = '" & CLng(Trim(GetText(vasID, iRow, colReceno))) & "' " & vbCrLf & _
'                              "and exam_code = '" & lsExamCode & "'"
'                    ElseIf Trim(GetText(vasID, iRow, colPSex)) = "F" Then
'                        SQL = "select refer_low_woman, refer_high_woman from totres " & vbCrLf & _
'                              "where request_date = '" & Trim(GetText(vasID, iRow, colReqDate)) & "' " & vbCrLf & _
'                              "and exam_no = '" & CLng(Trim(GetText(vasID, iRow, colSampleNo))) & "' " & vbCrLf & _
'                              "and exam_code = '" & lsExamCode & "'"
'                    End If
'
'                    res = db_select_Col(gServer, SQL)
                    
                    
                    
                    If lsResChk = "1" Then
                        If IsNumeric(lsResult) Then
                            If IsNumeric(gReadBuf(3)) = True And IsNumeric(gReadBuf(4)) = True Then
                                If CCur(lsResult) < gReadBuf(3) Then
                                    lsResult = "음성(" & lsResult & ")"
                                ElseIf CCur(lsResult) >= gReadBuf(3) And CCur(lsResult) < gReadBuf(4) Then
                                    lsResult = "약양성(" & lsResult & ")"
                                ElseIf CCur(lsResult) >= gReadBuf(4) Then
                                    lsResult = "양성(" & lsResult & ")"
                                End If
                            End If
                        End If
                        
                    Else
                        lsRefer = ""
                        If IsNumeric(lsResult) Then
                            If gReadBuf(3) <> "" And gReadBuf(4) <> "" Then
                                If CCur(gReadBuf(3)) > CCur(lsResult) Then
                                    lsRefer = "L"
                                ElseIf CCur(gReadBuf(4)) < CCur(lsResult) Then
                                    lsRefer = "H"
                                End If
                                
                            End If
                        End If
                        
                        vasRes.SetText colRefFlag, lResRow, lsRefer
                        vasRes.SetText 8, lResRow, gReadBuf(3)
                        vasRes.SetText 9, lResRow, gReadBuf(4)
                        
                        If lsRefer = "H" Then
                            vasRes.Row = lResRow
                            vasRes.Col = colRefFlag
                            vasRes.ForeColor = RGB(205, 55, 0)
                            vasRes.Col = colResult
                            vasRes.ForeColor = RGB(205, 55, 0)
                        ElseIf lsRefer = "L" Then
                            vasRes.Row = lResRow
                            vasRes.Col = colRefFlag
                            vasRes.ForeColor = RGB(0, 55, 205)
                            vasRes.Col = colResult
                            vasRes.ForeColor = RGB(0, 55, 205)
                        End If
                    
                    End If
                    
'''                    '소수점 처리
'''                    If IsNumeric(lsPoint) = True And IsNumeric(lsResult) = True Then
'''                        If CInt(lsPoint) > 0 Then
'''                            sTmpStr = "#0."
'''                            For i = 1 To CInt(lsPoint)
'''                                sTmpStr = sTmpStr & "0"
'''                            Next i
'''                        Else
'''                            sTmpStr = "#0"
'''                        End If
'''
'''                        lsResult = Format(lsResult, sTmpStr)
'''                    End If
                    
'''                    SetText vasRes, lsID, lResRow, 2  '검체번호
                    vasRes.SetText colEquipCode, lResRow, lsEquipCode
                    vasRes.SetText colResult, lResRow, lsResult
                    SetText vasRes, lsEquipCode, lResRow, colEquipCode      '장비코드
                    SetText vasRes, lsExamCode, lResRow, colExamCode        '검사코드
                    SetText vasRes, lsExamName, lResRow, colExamName        '검사명
                    SetText vasRes, lsResult, lResRow, colResult            '검사결과
'''                    SetText vasRes, lsResult, lResRow, colResult1           '검사결과
                    
                    
                    '로컬
                    Save_Local_Chk iRow, lResRow, lsExamGubun
                    
                End If
            End If
        
                
        End If
    Next iArr

                    
    '수신중========================================================
    SetText vasID, "Result", iRow, colState
    SetBackColor vasID, iRow, iRow, 1, 1, 255, 250, 205
    '==============================================================
    
End Function

Private Sub Make_Order(asRow As Long)
'Order 만들고 전송하기
    Dim sRetOrder As String     'Order Text넣을 변수
    Dim sOrder As String

    Dim i As Integer
    Dim iRow As Long


    Dim llRow As Long
    Dim llRow_Order As Long

    Dim sBarcode As String      '검체번호
    Dim sPID As String
    Dim sReceNo As String       '접수번호
    Dim sRackNo As String
    Dim sPosNo As String
    Dim sORDT As String         '접수일자
    Dim sExamCode As String     '검사코드
    Dim sEquipCode As String    '장비코드
    Dim sOrderCode As String
    Dim sState As String


    Dim lsCurDate As String
    Dim lsSampleNo As String
    Dim lsType As String

    Dim x As Integer
    Dim S  As String
    Dim j As Integer
    Dim k As Integer

    Dim sCnt As String
    Dim lsEquipCodeYN As Integer

    On Error GoTo errorchk


    gOrderMessage = ""

    lsCurDate = Format(Date, "yyyymmdd") & Format(Time, "hhnnss")

    llRow_Order = 1
    iRow = asRow

    'Order 만들기================================================
    If IsNumeric(Trim(GetText(vasID, iRow, colSampleNo))) Then
        lsSampleNo = Trim(GetText(vasID, iRow, colSampleNo))
    Else
        lsSampleNo = "1"
    End If

    ClearSpread vasOrderBuf
'    ClearSpread vasOrder

    sRetOrder = ""

    sBarcode = ""
    sReceNo = ""

    glRow = 0
    llRow = 1

    sReceNo = Trim(GetText(vasID, iRow, colReceno))

    sPID = Trim(GetText(vasID, iRow, colPID))
    sORDT = Trim(GetText(vasID, iRow, colReceDate))

    sBarcode = ""
    sBarcode = Mid(sORDT, 3) & "-" & sReceNo

    '====================================================
    ClearSpread vasCode

    '검사코드, 검사항목코드 가져오기
'            SQL = " Select examcode " & vbCrLf & _
'                  " From pat_res " & vbCrLf & _
'                  " Where examdate = '" & Format(frmInterface.txtToday.Text, "yyyymmdd") & "' " & vbCrLf & _
'                  " And equipno = '" & gEquip & "' " & vbCrLf & _
'                  " And barcode = '" & Trim(sBarcode) & "' " & vbCrLf & _
'                  " And examcode in (" & gAllExam & ") "
'
'            res = db_select_Vas(gLocal, SQL, vasCode)

'            If res = 0 Then     'Server에서 가져오기

'        If cn_Server_Flag = True Then
'            DisConnect_Server
'        Else
'            Connect_Server
'        End If
'
    SQL = "select examcode, equipcode from pat_res where equipno = '" & gEquip & "' " & vbCrLf & _
         "and recedate = '" & sORDT & "' and receno = '" & sReceNo & "'"
    res = db_select_Vas(gLocal, SQL, vasCode)
    
    If res = -1 Then
        SaveQuery SQL
    End If
'            End If
    '====================================================

    'Order 생성
    sOCnt = 1

    sOrderCode = ""

    ClearSpread vasCode_1

    For i = 1 To vasCode.DataRowCnt
        sExamCode = Trim(GetText(vasCode, i, 1))

        '검사코드로 장비코드 불러오기
        sEquipCode = Trim(GetText(vasCode, i, 2))
        SetText vasCode, sEquipCode, i, 3
    
        lsEquipCodeYN = 0
        For x = 1 To vasCode_1.DataRowCnt
            If Trim(GetText(vasCode_1, x, 1)) = sEquipCode Then
                lsEquipCodeYN = 1
                Exit For
            End If
        Next
        If lsEquipCodeYN = 0 Then
            SetText vasCode_1, sEquipCode, vasCode_1.DataRowCnt + 1, 1
        End If


        If sEquipCode <> "" Then
            sOCnt = sOCnt + 1
            SetText vasList, sOCnt, iRow, 14
'                    sOrderCode = sOrderCode & "^^^" & sEquipCode & "^\"
        End If
    Next i

    For x = 1 To vasCode_1.DataRowCnt
        Select Case Trim(GetText(vasCode_1, x, 1))
        Case "961"
            If InStr(1, sOrderCode, "678") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "678" & "^\"
            End If
            
            If InStr(1, sOrderCode, "413") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "413" & "^\"
            End If
            
        Case "962"
            If InStr(1, sOrderCode, "798") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "798" & "^\"
            End If
            
            If InStr(1, sOrderCode, "435") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "435" & "^\"
            End If
            
            If InStr(1, sOrderCode, "781") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "781" & "^\"
            End If
            
        Case "963"
            If InStr(1, sOrderCode, "421") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "421" & "^\"
            End If
            
            If InStr(1, sOrderCode, "690") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "690" & "^\"
            End If
        Case "964"
            If InStr(1, sOrderCode, "294") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "294" & "^\"
            End If
            
            If InStr(1, sOrderCode, "18") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "18" & "^\"
            End If
            
        Case "965"
            If InStr(1, sOrderCode, "678") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "678" & "^\"
            End If
            
            If InStr(1, sOrderCode, "413") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "413" & "^\"
            End If
            
        Case "966"
            If InStr(1, sOrderCode, "798") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "798" & "^\"
            End If
            
            If InStr(1, sOrderCode, "435") > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & "435" & "^\"
            End If
            
        Case Else
            If InStr(1, sOrderCode, Trim(GetText(vasCode_1, x, 1))) > 0 Then
            Else
                sOrderCode = sOrderCode & "^^^" & Trim(GetText(vasCode_1, x, 1)) & "^\"
            End If
        End Select
    Next

    ClearSpread vasCode_1

    If Trim(sOrderCode) <> "" Then
        sOrderCode = Left(sOrderCode, Len(sOrderCode) - 1)
    End If



    sOrder = ""

    vasList.Row = iRow
    vasList.Col = 4
    lsType = vasList.TypeComboBoxCurSel + 1
    If lsType < 1 Then
        lsType = 1
    End If

    'Order 전송하기==============================================

    If chkEM.Value = 1 Then
        sRackNo = Format(txtRack, "00#")
        sRackNo = "4" & Right(sRackNo, 3)
        sPosNo = txtPos

        If sOCnt > 0 Then
           sOrder = "H|\^&|||host^2|||||cobas6000|TSDWN^BATCH|P|1" & chrCR & _
                    "P|1|||||||U||||||^" & chrCR & _
                    "O|1|" & SetSpace(sBarcode, 22, 1) & "|^" & sRackNo & "^" & sPosNo & "^^S1^SC" & "|" & sOrderCode & "|S||" & lsCurDate & "||||N||||1||||||||||O" & chrCR & _
                    "L|1|N" & chrCR

            SetText vasID, sRackNo, iRow, colRack
            SetText vasID, sPosNo, iRow, colPos
            
           sPosNo = CLng(sPosNo) + 1
           If CLng(sPosNo) > 5 Then
                sPosNo = 1
                sRackNo = CLng(sRackNo) + 1

            End If
'                    If Trim(GetText(vasList, iRow, 5)) = "" Then
'                       SetText vasList, lsSampleNo, iRow, 5
'                   End If
'
'                   lsSampleNo = CLng(lsSampleNo) + 1
'                   txtSNo = lsSampleNo
           txtRack = sRackNo
           txtPos = sPosNo
       Else

           sOrder = "H|\^&|||host^2|||||cobas6000|TSDWN^BATCH|P|1" & chrCR & _
                    "P|1|||||||U||||||^" & chrCR & _
                    "O|1|" & SetSpace(sBarcode, 22, 1) & "|^" & sRackNo & "^" & sPosNo & "^^S1^SC" & "|^^^ALL|S||" & lsCurDate & "||||N||||1||||||||||O" & chrCR & _
                    "L|1|N" & chrCR
       End If
    Else
       If sOCnt > 0 Then
           sOrder = "H|\^&|||host^2|||||cobas6000|TSDWN^BATCH|P|1" & chrCR & _
                    "P|1|||||||U||||||^" & chrCR & _
                    "O|1|" & SetSpace(sBarcode, 22, 1) & "|" & lsSampleNo & "^^^^S1^SC" & "|" & sOrderCode & "|R||" & lsCurDate & "||||N||||1||||||||||O" & chrCR & _
                    "L|1|N" & chrCR

            
'            SetText vasID, lsSampleNo, iRow, colSampleNo
            
'           lsSampleNo = CLng(lsSampleNo) + 1
'           txtSNo = lsSampleNo
       Else

           sOrder = "H|\^&|||host^2|||||cobas6000|TSDWN^BATCH|P|1" & chrCR & _
                    "P|1|||||||U||||||^" & chrCR & _
                    "O|1|" & SetSpace(sBarcode, 22, 1) & "|" & lsSampleNo & "^^^^S1^SC" & "|^^^ALL|R||" & lsCurDate & "||||N||||1||||||||||O" & chrCR & _
                    "L|1|N" & chrCR
       End If
    End If

    llRow_Order = vasOrder.DataRowCnt + 1
    If llRow_Order > frmInterface.vasOrder.MaxRows Then
        frmInterface.vasOrder.MaxRows = llRow_Order
    End If
    
    SetText frmInterface.vasOrder, sOrder, llRow_Order, 1

'    llRow_Order = llRow_Order + 1
    

    Exit Sub

errorchk:
    MsgBox "전송중 에러가 있습니다. 확인"
    Me.MousePointer = 0
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
    
    If sEquipRes = "neg" Then
        sResult = "-"
    ElseIf sEquipRes = "norm" Then
        sResult = "+/-"
    ElseIf sEquipRes = "pos" Then
        sResult = "+"
    Else
        sResult = sEquipRes
    End If
    
        
'    If IsNumeric(sEquipRes) = False Then
'        Exit Function
'    End If
'
'    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' "
'    res = db_select_Col(gLocal, SQL)
'
'    If IsNumeric(gReadBuf(0)) = True Then
'        sPoint = CInt(gReadBuf(0))
'        sResType = ""
'        For i = 0 To sPoint
'            If i = 0 Then
'                sResType = "#0"
'            ElseIf i = 1 Then
'                sResType = sResType & ".0"
'            Else
'                sResType = sResType & "0"
'            End If
'        Next
'
'        sResult = Format(sEquipRes, sResType)
'    Else
'        sResult = sEquipRes
'    End If
'
'    If IsNumeric(gReadBuf(1)) = True Then
'        sLVal = gReadBuf(1)
'        If CCur(sLVal) > CCur(sEquipRes) Then
'            sResFlag = "<"
'        End If
'    End If
'
'    If IsNumeric(gReadBuf(2)) = True Then
'        sHVal = gReadBuf(2)
'        If CCur(sHVal) < CCur(sEquipRes) Then
'            sResFlag = ">"
'        End If
'    End If
'
'    sResult = sResFlag & sResult
    SetResult = sResult
    
End Function

Function Save_Local_Chk(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As Boolean)
    Dim sCnt As String
    Dim sExamDate As String
    Dim RCnt As Integer
    Dim OCnt As Integer
    
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    If asSend = True Then
        SQL = "update pat_res set result = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
              "                   refflag = '" & Trim(GetText(vasRes, asRow2, colRefFlag)) & "', sendflag = '2' " & vbCrLf & _
              "where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and recedate = '" & Trim(GetText(vasID, asRow1, colReceDate)) & "' " & vbCrLf & _
              "  and receno = '" & Trim(GetText(vasID, asRow1, colReceno)) & "' " & vbCrLf & _
              "  and examcode = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "' " & vbCrLf & _
              "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'"
        res = SendQuery(gLocal, SQL)
    Else
        SQL = "select sampleno from pat_res " & vbCrLf & _
              "where equipno = '" & gEquip & "' and examdate = '" & Format(Date, "yyyymmdd") & "' and sampleno = '" & Trim(GetText(vasID, asRow1, colSampleNo)) & "'  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'"
        res = db_select_Col(gLocal, SQL)
        
        If res > 0 Then
            SQL = "update pat_res set result = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
                  "                   refflag = '" & Trim(GetText(vasRes, asRow2, colRefFlag)) & "' " & vbCrLf & _
                  "where equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and examdate = '" & Format(Date, "yyyymmdd") & "' " & vbCrLf & _
                  "  and sampleno = '" & Trim(GetText(vasID, asRow1, colSampleNo)) & "' and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'"
            res = SendQuery(gLocal, SQL)
        
        Else
            SQL = "insert into pat_res(equipno, examdate, sampleno, equipcode, result, sendflag, examname, hospital) " & vbCrLf & _
                  "values('" & gEquip & "', '" & Format(Date, "yyyymmdd") & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasID, asRow1, colSampleNo)) & "', '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', " & vbCrLf & _
                  "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', '0','" & Trim(GetText(vasRes, asRow2, colExamName)) & "', '미지정' )"
            res = SendQuery(gLocal, SQL)
            
        End If
    End If
    
End Function



Function Save_Local_One(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    Dim RCnt As Integer
    Dim OCnt As Integer
    
    sExamDate = Format(dtpToday, "yyyymmdd")

    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND receno = '" & Trim(GetText(vasID, asRow1, colReceno)) & "' " & vbCrLf & _
          "  and sampleno = '" & Trim(GetText(vasID, asRow1, colSampleNo)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'"
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "INSERT INTO PAT_RES(EQUIPNO, RECENO, DISKNO, " & vbCrLf & _
          "POSNO, PID, PNAME, " & vbCrLf & _
          "PJUMIN, PSEX, PAGE, " & vbCrLf & _
          "EXAMDATE, EQUIPCODE, EXAMCODE, " & vbCrLf & _
          "SEQNO, RESULT, EXAMNAME, SENDFLAG,SAMPLENO,examgubun,subcode,hospital) " & vbCrLf & _
          "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, asRow1, colReceno)) & "', '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colPos)) & "', '" & Trim(GetText(vasID, asRow1, colPID)) & "', '" & Trim(GetText(vasID, asRow1, colPName)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasID, asRow1, colPJumin)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " & 0 & ", " & vbCrLf & _
          "'" & Trim(sExamDate) & "', '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', '" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
          "'" & asSend & "','" & Trim(GetText(vasID, asRow1, colSampleNo)) & "','" & Trim(GetText(vasID, asRow1, colExamDate)) & "','" & Trim(GetText(vasRes, asRow2, colSubCode)) & "','미지정')"
    res = SendQuery(gLocal, SQL)

    
End Function

Function Insert_Data(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim lRow    As Integer
    Dim lsID    As String
    Dim vasRow  As Integer
    Dim lsDBDate As String
    Dim sRef As String
    Dim sHL As String
    Dim SubCode As String
    
    Insert_Data = -1
    lRow = argSpcRow
    
    If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Function

    lsID = Trim(GetText(vasID, lRow, colReceno))
    lsDBDate = Trim(GetText(vasID, lRow, colReceDate))
    If lsID = "" Then Exit Function

    ClearSpread vasTemp
    
    SQL = "SELECT EQUIPCODE, EXAMCODE, SUBCODE,EXAMNAME, RESULT, examgubun  " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND examdate = '" & Format(Trim(dtpToday.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND RECENO = '" & lsID & "' " & vbCrLf & _
          "  AND PID = '" & GetText(vasID, lRow, colPID) & "' " & vbCrLf & _
          "  AND SAMPLENO = '" & GetText(vasID, lRow, colSampleNo) & "'"
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    For vasRow = 1 To vasTemp.DataRowCnt
    
        Select Case Trim(GetText(vasTemp, vasRow, 6))
        
        Case dpGumjin1
            If Trim(GetText(vasID, lRow, colPSex)) = "M" Then
                SQL = "select refer_low_man, refer_high_man from totres " & vbCrLf & _
                      "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                      "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                      "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            ElseIf Trim(GetText(vasID, lRow, colPSex)) = "F" Then
                SQL = "select refer_low_woman, refer_high_woman from totres " & vbCrLf & _
                      "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                      "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                      "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            End If
            
            res = db_select_Col(gServer, SQL)
            sRef = ""
            sHL = ""
            If IsNumeric(Trim(GetText(vasTemp, vasRow, 5))) Then
                If gReadBuf(0) <> "" And gReadBuf(1) <> "" Then
                    If CCur(gReadBuf(0)) > CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "L"
                    ElseIf CCur(gReadBuf(1)) < CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "H"
                    End If
                    
                End If
            End If
            
            SQL = "update totres " & vbCrLf & _
                  "set result_value = '" & Trim(GetText(vasTemp, vasRow, 5)) & "', " & vbCrLf & _
                  " result_decision = '" & sRef & "' " & vbCrLf & _
                  "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                  "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                  "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                'db_RollBack gServer
                Exit Function
            End If
        Case dpOCS
        
            If Trim(GetText(vasID, lRow, colPSex)) = "M" Then
                SQL = "select RowValueMan, HighValueMan from TB_진료검사 " & vbCrLf & _
                      " Where 년 = '" & Mid(lsDBDate, 1, 4) & "' " & vbCrLf & _
                      " And 월 = '" & Mid(lsDBDate, 5, 2) & "' " & vbCrLf & _
                      " And 일 = '" & Mid(lsDBDate, 7, 2) & "' " & vbCrLf & _
                      " And 챠트번호 = '" & Trim(GetText(vasID, lRow, colReceno)) & "' " & vbCrLf & _
                      " And 검사코드 = '" & Trim(GetText(vasTemp, vasRow, 2)) & "' "
                      
            ElseIf Trim(GetText(vasID, lRow, colPSex)) = "F" Then
                SQL = "select RowValueWoman, HighValueWoman from TB_진료검사 " & vbCrLf & _
                      " Where 년 = '" & Mid(lsDBDate, 1, 4) & "' " & vbCrLf & _
                      " And 월 = '" & Mid(lsDBDate, 5, 2) & "' " & vbCrLf & _
                      " And 일 = '" & Mid(lsDBDate, 7, 2) & "' " & vbCrLf & _
                      " And 챠트번호 = '" & Trim(GetText(vasID, lRow, colReceno)) & "' " & vbCrLf & _
                      " And 검사코드 = '" & Trim(GetText(vasTemp, vasRow, 2)) & "' "
            End If

            res = db_select_Col(gServer_OCS, SQL)
            sRef = ""
            sHL = ""
            If IsNumeric(Trim(GetText(vasTemp, vasRow, 5))) Then
                If gReadBuf(0) <> "" And gReadBuf(1) <> "" Then
                    If CCur(gReadBuf(0)) > CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "L"
                    ElseIf CCur(gReadBuf(1)) < CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "H"
                    End If

                End If
            End If
            
            SQL = "Update TB_진료검사 " & vbCrLf & _
                  "Set " & vbCrLf & _
                  "    결과 = '" & Trim(GetText(vasTemp, vasRow, 5)) & "' ," & vbCrLf & _
                  "    수정일 =  getdate() ," & vbCrLf & _
                  "    값 = '" & Trim(sRef) & "', " & vbCrLf & _
                  "    상태 = '1' "
            SQL = SQL & CR & _
              " Where 년 = '" & Mid(lsDBDate, 1, 4) & "' " & vbCrLf & _
              " And 월 = '" & Mid(lsDBDate, 5, 2) & "' " & vbCrLf & _
              " And 일 = '" & Mid(lsDBDate, 7, 2) & "' " & vbCrLf & _
              " And 챠트번호 = '" & Trim(GetText(vasID, lRow, colReceno)) & "' " & vbCrLf & _
              " AND 상태 = '0' " & vbCrLf & _
              " And 검사코드 = '" & Trim(GetText(vasTemp, vasRow, 2)) & "' " & vbCrLf & _
              " And 검사종류 = '" & Format(Trim(GetText(vasTemp, vasRow, 3)), "0#") & "' "
    
            res = SendQuery(gServer_OCS, SQL)
            If res < 0 Then
                SaveQuery SQL
    '            db_RollBack gServer
                Exit Function
            End If
            
                        
            
        Case dpGumjin2
            If Trim(GetText(vasID, lRow, colPSex)) = "M" Then
                SQL = "select refer_low_man, refer_high_man from twoexam " & vbCrLf & _
                      "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                      "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                      "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            ElseIf Trim(GetText(vasID, lRow, colPSex)) = "F" Then
                SQL = "select refer_low_woman, refer_high_woman from twoexam " & vbCrLf & _
                      "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                      "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                      "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            End If

            res = db_select_Col(gServer, SQL)
            sRef = ""
            sHL = ""
            If IsNumeric(Trim(GetText(vasTemp, vasRow, 5))) Then
                If gReadBuf(0) <> "" And gReadBuf(1) <> "" Then
                    If CCur(gReadBuf(0)) > CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "L"
                    ElseIf CCur(gReadBuf(1)) < CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "H"
                    End If

                End If
            End If

            If IsNumeric(Trim(GetText(vasTemp, vasRow, 5))) = False Then
            gReadBuf(0) = ""
            SQL = "select exam_subcode from examref " & vbCrLf & _
                    "where ref_value = '" & Trim(GetText(vasTemp, vasRow, 5)) & "' " & vbCrLf & _
                    "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            res = db_select_Col(gServer, SQL)
            SubCode = gReadBuf(0)
            End If

            SQL = "update twoexam " & vbCrLf & _
                  "set result_value = '" & Trim(GetText(vasTemp, vasRow, 5)) & "', " & vbCrLf & _
                  " result_decision = '" & sRef & "' " & vbCrLf & _
                  "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                  "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                  "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
                        If Trim(GetText(vasTemp, vasRow, 3)) <> "" Then
            SQL = SQL & CR & _
                 " And sub_code = '" & SubCode & "' "
            End If

            res = SendQuery(gServer, SQL)

            If res = -1 Then
                'db_RollBack gServer
                Exit Function
            End If
        End Select
        
    Next vasRow
    
'    db_Commit gServer
    
    Insert_Data = 1
    
    Exit Function

End Function

Function Insert_Data_R(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim lRow    As Integer
    Dim lsID    As String
    Dim vasRow  As Integer
    Dim lsDBDate As String
    Dim sRef As String
    Dim sHL As String
    Dim SubCode As String


    Insert_Data_R = -1
    lRow = argSpcRow

    If lRow < 1 Or lRow > vasRID.DataRowCnt Then Exit Function

    lsID = Trim(GetText(vasRID, lRow, colReceno))
    lsDBDate = Trim(GetText(vasRID, lRow, colReceDate))
    If lsID = "" Then Exit Function

    ClearSpread vasTemp

    SQL = "SELECT EQUIPCODE, EXAMCODE, SUBCODE,EXAMNAME, RESULT, examgubun  " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND RECENO = '" & lsID & "' " & vbCrLf & _
          "  AND PID = '" & GetText(vasRID, lRow, colPID) & "' " & vbCrLf & _
          "  AND SAMPLENO = '" & GetText(vasRID, lRow, colSampleNo) & "'"
    res = db_select_Vas(gLocal, SQL, vasTemp)

'    db_BeginTran
    For vasRow = 1 To vasTemp.DataRowCnt
    
        Select Case Trim(GetText(vasTemp, vasRow, 6))
        
        Case dpGumjin1
            If Trim(GetText(vasRID, lRow, colPSex)) = "M" Then
                SQL = "select refer_low_man, refer_high_man from totres " & vbCrLf & _
                      "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                      "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                      "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            ElseIf Trim(GetText(vasRID, lRow, colPSex)) = "F" Then
                SQL = "select refer_low_woman, refer_high_woman from totres " & vbCrLf & _
                      "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                      "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                      "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            End If
            
            res = db_select_Col(gServer, SQL)
            sRef = ""
            sHL = ""
            If IsNumeric(Trim(GetText(vasTemp, vasRow, 5))) Then
                If gReadBuf(0) <> "" And gReadBuf(1) <> "" Then
                    If CCur(gReadBuf(0)) > CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "L"
                    ElseIf CCur(gReadBuf(1)) < CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "H"
                    End If
                    
                End If
            End If
            
            If IsNumeric(Trim(GetText(vasTemp, vasRow, 5))) = False Then
            gReadBuf(0) = ""
            SQL = "select exam_subcode from examref " & vbCrLf & _
                    "where ref_value = '" & Trim(GetText(vasTemp, vasRow, 5)) & "' " & vbCrLf & _
                    "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            res = db_select_Col(gServer, SQL)
            SubCode = gReadBuf(0)
            End If
            
            SQL = "update totres " & vbCrLf & _
                  "set result_value = '" & Trim(GetText(vasTemp, vasRow, 5)) & "', " & vbCrLf & _
                  " result_decision = '" & sRef & "' " & vbCrLf & _
                  "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                  "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                  "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
                        If Trim(GetText(vasTemp, vasRow, 3)) <> "" Then
            SQL = SQL & CR & _
                 " And sub_code = '" & SubCode & "' "
            End If
            
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                'db_RollBack gServer
                Exit Function
            End If
        Case dpOCS
        
            If Trim(GetText(vasRID, lRow, colPSex)) = "M" Then
                SQL = "select RowValueMan, HighValueMan from TB_진료검사 " & vbCrLf & _
                      " Where 년 = '" & Mid(lsDBDate, 1, 4) & "' " & vbCrLf & _
                      " And 월 = '" & Mid(lsDBDate, 5, 2) & "' " & vbCrLf & _
                      " And 일 = '" & Mid(lsDBDate, 7, 2) & "' " & vbCrLf & _
                      " And 챠트번호 = '" & Trim(GetText(vasRID, lRow, colReceno)) & "' " & vbCrLf & _
                      " And 검사코드 = '" & Trim(GetText(vasTemp, vasRow, 2)) & "' "
                      
            ElseIf Trim(GetText(vasRID, lRow, colPSex)) = "F" Then
                SQL = "select RowValueWoman, HighValueWoman from TB_진료검사 " & vbCrLf & _
                      " Where 년 = '" & Mid(lsDBDate, 1, 4) & "' " & vbCrLf & _
                      " And 월 = '" & Mid(lsDBDate, 5, 2) & "' " & vbCrLf & _
                      " And 일 = '" & Mid(lsDBDate, 7, 2) & "' " & vbCrLf & _
                      " And 챠트번호 = '" & Trim(GetText(vasRID, lRow, colReceno)) & "' " & vbCrLf & _
                      " And 검사코드 = '" & Trim(GetText(vasTemp, vasRow, 2)) & "' "
            End If

            res = db_select_Col(gServer_OCS, SQL)
            sRef = ""
            sHL = ""
            If IsNumeric(Trim(GetText(vasTemp, vasRow, 5))) Then
                If gReadBuf(0) <> "" And gReadBuf(1) <> "" Then
                    If CCur(gReadBuf(0)) > CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "L"
                    ElseIf CCur(gReadBuf(1)) < CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "H"
                    End If

                End If
            End If
            
            SQL = "Update TB_진료검사 " & vbCrLf & _
                  "Set " & vbCrLf & _
                  "    결과 = '" & Trim(GetText(vasTemp, vasRow, 5)) & "' ," & vbCrLf & _
                  "    수정일 =  getdate() ," & vbCrLf & _
                  "    값 = '" & Trim(sRef) & "', " & vbCrLf & _
                  "    상태 = '1' "
            SQL = SQL & CR & _
              " Where 년 = '" & Mid(lsDBDate, 1, 4) & "' " & vbCrLf & _
              " And 월 = '" & Mid(lsDBDate, 5, 2) & "' " & vbCrLf & _
              " And 일 = '" & Mid(lsDBDate, 7, 2) & "' " & vbCrLf & _
              " And 챠트번호 = '" & Trim(GetText(vasRID, lRow, colReceno)) & "' " & vbCrLf & _
              " And 검사코드 = '" & Trim(GetText(vasTemp, vasRow, 2)) & "' " & vbCrLf & _
              " and 상태 = '0' " & vbCrLf & _
              " And 검사종류 = '" & Format(Trim(GetText(vasTemp, vasRow, 3)), "0#") & "' "

    
            res = SendQuery(gServer_OCS, SQL)
            If res < 0 Then
                SaveQuery SQL
    '            db_RollBack gServer
                Exit Function
            End If
            
        Case dpGumjin2
            If Trim(GetText(vasRID, lRow, colPSex)) = "M" Then
                SQL = "select refer_low_man, refer_high_man from twoexam " & vbCrLf & _
                      "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                      "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                      "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            ElseIf Trim(GetText(vasRID, lRow, colPSex)) = "F" Then
                SQL = "select refer_low_woman, refer_high_woman from twoexam " & vbCrLf & _
                      "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                      "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                      "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            End If
            
            res = db_select_Col(gServer, SQL)
            sRef = ""
            sHL = ""
            If IsNumeric(Trim(GetText(vasTemp, vasRow, 5))) Then
                If gReadBuf(0) <> "" And gReadBuf(1) <> "" Then
                    If CCur(gReadBuf(0)) > CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "L"
                    ElseIf CCur(gReadBuf(1)) < CCur(Trim(GetText(vasTemp, vasRow, 5))) Then
                        sRef = "H"
                    End If
                    
                End If
            End If
            
            If IsNumeric(Trim(GetText(vasTemp, vasRow, 5))) = False Then
            gReadBuf(0) = ""
            SQL = "select exam_subcode from examref " & vbCrLf & _
                    "where ref_value = '" & Trim(GetText(vasTemp, vasRow, 5)) & "' " & vbCrLf & _
                    "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            res = db_select_Col(gServer, SQL)
            SubCode = gReadBuf(0)
            End If
            
            SQL = "update twoexam " & vbCrLf & _
                  "set result_value = '" & Trim(GetText(vasTemp, vasRow, 5)) & "', " & vbCrLf & _
                  " result_decision = '" & sRef & "' " & vbCrLf & _
                  "where request_date = '" & lsDBDate & "' " & vbCrLf & _
                  "and exam_no = '" & CLng(lsID) & "' " & vbCrLf & _
                  "and exam_code = '" & Trim(GetText(vasTemp, vasRow, 2)) & "'"
            If Trim(GetText(vasTemp, vasRow, 3)) <> "" Then
            SQL = SQL & CR & _
                 " And sub_code = '" & SubCode & "' "
            End If
            
            res = SendQuery(gServer, SQL)
            
            If res = -1 Then
                'db_RollBack gServer
                Exit Function
            End If
        End Select
        
    Next vasRow
'    db_Commit

    Insert_Data_R = 1

    Exit Function

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

Private Sub subdel_Click()
    Dim i As Integer
    Dim Response, Help
    
    Response = MsgBox("삭제하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!", Help, 100)
    
    If Response = vbYes Then
        For i = vasRID.DataRowCnt To 1 Step -1
            vasRID.Col = 1
            vasRID.Row = i
            If vasRID.Value = 1 Then
            
                SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                      "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
                      "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      "  AND receno = '" & Trim(GetText(vasRID, i, colReceno)) & "' " & vbCrLf & _
                      "  and sampleno = '" & Trim(GetText(vasRID, i, colSampleNo)) & "' "
                res = SendQuery(gLocal, SQL)
                
                vasRID.Col = 1
                vasRID.Row = i
                vasRID.Value = 0
        
                DeleteRow vasRID, i, i

            End If
        Next i
    End If
End Sub

Private Sub Timer1_Timer()
    dtpToday = Date
End Sub

Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim iRow1 As Long
    Dim iRow2 As Long
    Dim i As Long
    
    iRow1 = BlockRow
    iRow2 = BlockRow2
    
    For i = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = i
        vasID.Value = 0
    Next
    
    For i = iRow1 To iRow2
        vasID.Col = 1
        vasID.Row = i
        vasID.Value = 1
    Next i
End Sub

'Private Sub Picture1_Click()
'    frmUser.Show 0
'
'End Sub

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colReceno))
    
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, SUBCODE, EXAMNAME, RESULT, seqno " & vbCrLf & _
          "FROM PAT_RES" & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND isnull(RECENO, '') = '" & lsID & "'  " & vbCrLf & _
          "AND isnull(recedate, '') = '" & Trim(GetText(vasID, Row, colReceDate)) & "' " & vbCrLf & _
          "AND examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' " & vbCrLf & _
          "AND sampleno = '" & Trim(GetText(vasID, Row, colSampleNo)) & "' " & vbCrLf & _
          "GROUP BY  seqno , EQUIPCODE, EXAMCODE, EXAMNAME, subcode, RESULT "
            
    res = db_select_Vas(gLocal, SQL, vasRes)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasRes.MaxRows = vasRes.DataRowCnt
End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
 Dim i As Long
    Dim j As Long
    Dim x As Long
    Dim intListRow As Long
    Dim intSampleNo As Long
    Dim chServer
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSubCode As String
    Dim lsSeqNo As String
    Dim lsEquipCode As String
    Dim result As String
    Dim sHospital As String
    
    chServer = cmbPart.ListIndex
    
    For i = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = i
        If vasList.Value = 1 Then
  
            SetText vasID, Trim(GetText(vasList, i, 2)), vasID.Row, colReceDate
            SetText vasID, Trim(GetText(vasList, i, 4)), vasID.Row, colReceno
            SetText vasID, Trim(GetText(vasList, i, 5)), vasID.Row, colPID
            SetText vasID, Trim(GetText(vasList, i, 6)), vasID.Row, colPName
            SetText vasID, Trim(GetText(vasList, i, 7)), vasID.Row, colPSex
            SetText vasID, Trim(GetText(vasList, i, 8)), vasID.Row, colPAge
            SetText vasID, Trim(GetText(vasList, i, 9)), vasID.Row, colPJumin
            SetText vasID, Trim(GetText(vasList, i, 10)), vasID.Row, colExamDate

            ClearSpread vasTemp
            
            Select Case chServer
            Case dpGumjin1
                SQL = "select exam_code, '' from totres " & vbCrLf & _
                      "where request_date = '" & Trim(GetText(vasID, vasID.Row, colReceDate)) & "' " & vbCrLf & _
                      "and exam_no = '" & Trim(GetText(vasID, vasID.Row, colReceno)) & "' and exam_code in (" & gAllExam & ") "
                res = db_select_Vas(gServer, SQL, vasTemp)
                
            Case dpOCS
                SQL = "select 검사코드, 검사종류 from tb_진료검사 " & vbCrLf & _
                      "where 년 = '" & Mid(Trim(GetText(vasID, vasID.Row, colReceDate)), 1, 4) & "' " & vbCrLf & _
                      "and 월 = '" & Mid(Trim(GetText(vasID, vasID.Row, colReceDate)), 5, 2) & "' " & vbCrLf & _
                      "and 일 = '" & Mid(Trim(GetText(vasID, vasID.Row, colReceDate)), 7, 2) & "' and 챠트번호 = '" & Trim(GetText(vasID, vasID.Row, colReceno)) & "' and 검사코드 in (" & gAllExam_Ocs & ") and 오더일련번호 > '0' "
                res = db_select_Vas(gServer_OCS, SQL, vasTemp)
            Case dpGumjin2
                SQL = "select exam_code, '' from twoexam " & vbCrLf & _
                      "where request_date = '" & Trim(GetText(vasID, vasID.Row, colReceDate)) & "' " & vbCrLf & _
                      "and exam_no = '" & Trim(GetText(vasID, vasID.Row, colReceno)) & "' and exam_code in (" & gAllExam & ") "
                res = db_select_Vas(gServer, SQL, vasTemp)
            End Select
            
            
            For x = 1 To vasTemp.DataRowCnt
                lsExamCode = Trim(GetText(vasTemp, x, 1))
                lsSubCode = Trim(GetText(vasTemp, x, 2))
                               
                SQL = "select examcode, subcode, equipcode, seqno, examname from equipexam " & vbCrLf & _
                      "where equipno = '" & gEquip & "' and examcode = '" & lsExamCode & "' and isnull(subcode, '') = '" & lsSubCode & "'"
                res = db_select_Col(gLocal, SQL)
                
                If res > 0 Then
                
                lsEquipCode = Trim(gReadBuf(2))
                lsSeqNo = Trim(gReadBuf(3))
                lsExamName = Trim(gReadBuf(4))
                
                If lsEquipCode = "" Then
                Else
                
                   gReadBuf(0) = ""
                   gReadBuf(1) = ""
                   
                   SQL = "select result,hospital from pat_res " & vbCrLf & _
                         "where equipno = '" & gEquip & "' and sampleno = '" & Trim(GetText(vasID, vasID.Row, 3)) & "' " & vbCrLf & _
                         "and equipcode = '" & lsEquipCode & "' and examdate = '" & Format(dtpToday.Value, "yyyymmdd") & "' AND SENDFLAG = '0' "
                   res = db_select_Col(gLocal, SQL)
                   
                   result = Trim(gReadBuf(0))
                   sHospital = Trim(gReadBuf(1))
                   
                   If chServer = dpGumjin1 Or chServer = dpGumjin2 Then
                      Select Case result
                          Case "trace"
                           result = "약양성"
                          Case "norm"
                           result = "음성"
                      End Select
                   End If
                   
                   If res > 0 Then
            
                    SQL = "insert into pat_res(equipno, recedate, sampleno, receno, pid, " & vbCrLf & _
                          "pname, psex, page, pjumin, sendflag, " & vbCrLf & _
                          "examgubun,examdate, examcode, subcode, examname, equipcode, seqno,RESULT,hospital) " & vbCrLf & _
                          "values('" & gEquip & "', '" & Trim(GetText(vasList, i, 2)) & "', '" & Trim(GetText(vasID, vasID.Row, 3)) & "', " & vbCrLf & _
                          "'" & Trim(GetText(vasList, i, 4)) & "', '" & Trim(GetText(vasList, i, 5)) & "', '" & Trim(GetText(vasList, i, 6)) & "', " & vbCrLf & _
                          "'" & Trim(GetText(vasList, i, 7)) & "', '" & Trim(GetText(vasList, i, 8)) & "', '" & Trim(GetText(vasList, i, 9)) & "', " & vbCrLf & _
                          "'2', '" & chServer & "','" & Format(dtpToday, "YYYYMMDD") & "', '" & lsExamCode & "', " & vbCrLf & _
                          "'" & lsSubCode & "', '" & lsExamName & "', '" & lsEquipCode & "', '" & lsSeqNo & "','" & result & "','" & sHospital & "')"
                    res = SendQuery(gLocal, SQL)
                 End If
                End If
               End If
            Next x
            vasID.Row = vasID.Row + 1
        End If
    Next i
    
    For i = vasList.DataRowCnt To 1 Step -1
        vasList.Col = 1
        vasList.Row = i
        If vasList.Value = 1 Then
            DeleteRow vasList, i, i
        End If
        
    Next i
    
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    
    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If
        
        lsID = Trim(GetText(vasID, iRow, colReceno))
        
'        If Trim(GetText(vasID, iRow, colPJumin)) = "F" Then
'            If MsgBox("해당 QC 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'                Exit Sub
'            End If
'
'            lsTime = Trim(GetText(vasID, iRow, colPID))
'            If Len(lsTime) = 4 Then
'            Else
'                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
'            End If
'
'            SQL = "Delete From qc_res a " & vbCrLf & _
'                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
'                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
'                  "  and a.levelname = '" & lsID & "' "
'            res = SendQuery(gLocal, SQL)
'
'            Exit Sub
'        End If
            
        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
            
        SQL = " DELETE FROM PAT_RES " & vbCrLf & _
              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
              " AND BARCODE = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0
    End If
End Sub

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
            
        vasID_Click colReceno, lRow
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

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    PopupMenu Mnlist2
End Sub

Private Sub vasList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim iRow1 As Long
    Dim iRow2 As Long
    Dim i As Long
    
    iRow1 = BlockRow
    iRow2 = BlockRow2
    
    For i = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = i
        vasList.Value = 0
    Next
    
    For i = iRow1 To iRow2
        vasList.Col = 1
        vasList.Row = i
        vasList.Value = 1
    Next i
    
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim iRow As Long
    Dim iCol As Long
    Dim iExamType As String
    Dim lsReceDate As String
    Dim lsReceNo As String
    Dim i As Long
    Dim j As Long
    
    iRow = Row
    iCol = Col
    
    ClearSpread vasRes
    
    iExamType = Trim(GetText(vasList, iRow, 10))
    
    Select Case iExamType
    Case dpGumjin1
        SQL = "select exam_code, '' from totres " & vbCrLf & _
              "where request_date = '" & Trim(GetText(vasList, iRow, 2)) & "' and exam_no = '" & Trim(GetText(vasList, iRow, 4)) & "' and exam_code in (" & gAllExam & ")"
        res = db_select_Vas(gServer, SQL, vasRes, 1, 2)
    Case dpOCS
        SQL = "select 검사코드, 검사종류 from tb_진료검사 " & vbCrLf & _
              "where 년 + 월 + 일 = '" & Trim(GetText(vasList, iRow, 2)) & "' and 챠트번호 = '" & Trim(GetText(vasList, iRow, 4)) & "' and 오더일련번호 > '0' and 검사코드 in (" & gAllExam_Ocs & ")"
        res = db_select_Vas(gServer_OCS, SQL, vasRes, 1, 2)
    Case dpGumjin2
        SQL = "selct exam_code, '' from twores " & vbCrLf & _
              "where request_date = '" & Trim(GetText(vasList, iRow, 2)) & "' and exam_no = '" & Trim(GetText(vasList, iRow, 4)) & "' and exam_code in (" & gAllExam & ")"
        res = db_select_Vas(gServer, SQL, vasRes, 1, 2)
        
    End Select
    
    ClearSpread vasExamTemp
    
    SQL = "select examcode, subcode, examname from equipexam where equipno = '" & gEquip & "'"
    res = db_select_Vas(gLocal, SQL, vasExamTemp)
    
    For i = 1 To vasRes.DataRowCnt
        For j = 1 To vasExamTemp.DataRowCnt
            If Trim(GetText(vasRes, i, 2)) = Trim(GetText(vasExamTemp, j, 1)) And Trim(GetText(vasRes, i, 3)) = Trim(GetText(vasExamTemp, j, 2)) Then
                SetText vasRes, Trim(GetText(vasExamTemp, j, 3)), i, colExamName
                Exit For
            End If
        Next j
    Next i
    
    vasRes.RowHeight(-1) = 13
End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim iCol As Integer
    Dim iRow As Integer
    Dim iExamCnt As Long
    
    
    iCol = vasList.ActiveCol
    iRow = vasList.ActiveRow
    
    If KeyCode = 13 Then
        
        If iCol = 3 Then
            iExamCnt = Trim(GetText(vasList, iRow, iCol))
            For i = iRow To vasList.DataRowCnt
                SetText vasList, iExamCnt, i, 3
                iExamCnt = iExamCnt + 1
            Next
        End If
    End If
    
End Sub

Private Sub vasList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    PopupMenu Mnlist
    
End Sub

Private Sub vasRID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim iRow1 As Long
    Dim iRow2 As Long
    Dim i As Long
    
    iRow1 = BlockRow
    iRow2 = BlockRow2
    
    For i = 1 To vasRID.DataRowCnt
        vasRID.Col = 1
        vasRID.Row = i
        vasRID.Value = 0
    Next
    
    For i = iRow1 To iRow2
        vasRID.Col = 1
        vasRID.Row = i
        vasRID.Value = 1
    Next i
End Sub

Private Sub vasRID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim sPNo As String
    Dim lsReceDate As String
    Dim lsReceNo As String
    Dim lsPID As String
    Dim intExamType
    Dim lsPJumin As String
    Dim i As Long
    
    
    If Row < 1 Or Row > vasRID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasRID, Row, colReceno))
    sPNo = Trim(GetText(vasRID, Row, colSampleNo))
    
    lsReceDate = Trim(GetText(vasRID, Row, colReceDate))
    lsReceNo = Trim(GetText(vasRID, Row, colReceno))
    lsPID = Trim(GetText(vasRID, Row, colPID))
    lsPJumin = Trim(GetText(vasRID, Row, colPJumin))
    
    
    intExamType = cmbServerType.ListIndex
       
    
    'Local에서 불러오기
    ClearSpread vasRRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
'    SQL = "SELECT EQUIPCODE, EXAMCODE, subcode,EXAMNAME, RESULT, '','',SEQNO " & vbCrLf & _
'          "FROM PAT_RES " & vbCrLf & _
'          "WHERE EQUIPNO = '" & gEquip & "' AND receno = '" & lsID & "' " & vbCrLf & _
'          "and sampleno = '" & sPNo & "' " & vbCrLf & _
'          "GROUP BY SEQNO, EQUIPCODE, subcode,EXAMCODE, EXAMNAME, RESULT "
    SQL = "SELECT EQUIPCODE, EXAMCODE, SUBCODE, EXAMNAME, RESULT, '', '', seqno " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND isnull(RECENO, '') = '" & lsID & "'  " & vbCrLf & _
          "AND isnull(recedate, '') = '" & Trim(GetText(vasRID, Row, colReceDate)) & "' " & vbCrLf & _
          "AND examdate = '" & Format(dtpExamDate, "yyyymmdd") & "' " & vbCrLf & _
          "AND sampleno = '" & Trim(GetText(vasRID, Row, colSampleNo)) & "' " & vbCrLf & _
          "GROUP BY seqno, EQUIPCODE, EXAMCODE, subcode, EXAMNAME, RESULT "
          
    
    res = db_select_Vas(gLocal, SQL, vasRRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    If optExamPart(0).Value = True Then
    
        For i = 1 To vasRRes.DataRowCnt
        
            Select Case intExamType
            Case dpGumjin1
                SQL = "select max(a.request_date), a.Result_Value from totres a, total b " & vbCrLf & _
                      "where a.request_date = b.request_date and a.exam_no = b.exam_no " & vbCrLf & _
                      "and a.request_date <> '" & lsReceDate & "' and isnull(a.result_value, '') <> '' " & vbCrLf & _
                      "and a.exam_no <> '" & lsReceNo & "' and b.personal_id = '" & lsPJumin & "' and a.exam_code = '" & Trim(GetText(vasRRes, i, colExamCode)) & "' " & vbCrLf & _
                      "group by a.request_date, a.result_value  "
                res = db_select_Col(gServer, SQL)
                SetText vasRRes, Trim(gReadBuf(0)), i, 6
                SetText vasRRes, Trim(gReadBuf(1)), i, 7
            Case dpOCS
                SQL = "select max(수정일), 결과 from tb_진료검사 " & vbCrLf & _
                      "where 챠트번호 = '" & lsReceNo & "' and 검사코드 = '" & Trim(GetText(vasRRes, i, colExamCode)) & "' " & vbCrLf & _
                      "and 검사종류 = '" & Trim(GetText(vasRRes, i, colSubCode)) & "' " & vbCrLf & _
                      "and 상태 = '1' " & vbCrLf & _
                      "group by 수정일,결과 "
                res = db_select_Col(gServer_OCS, SQL)
                SetText vasRRes, Format(Trim(gReadBuf(0)), "yyyymmdd"), i, 6
                SetText vasRRes, Trim(gReadBuf(1)), i, 7
            
            Case dpGumjin2
                SQL = "select max(a.request_date), a.Result_Value from twoexam a, total b " & vbCrLf & _
                      "where a.request_date = b.request_date and a.exam_no = b.exam_no " & vbCrLf & _
                      "and a.request_date <> '" & lsReceDate & "' and isnull(a.result_value, '') <> '' " & vbCrLf & _
                      "and a.exam_no <> '" & lsReceNo & "' and b.personal_id = '" & lsPJumin & "' and a.exam_code = '" & Trim(GetText(vasRRes, i, colExamCode)) & "' " & vbCrLf & _
                      "group by a.request_date, a.result_value  "
                res = db_select_Col(gServer, SQL)
                SetText vasRRes, Trim(gReadBuf(0)), i, 6
                SetText vasRRes, Trim(gReadBuf(1)), i, 7
            End Select
        Next
    
    End If
    
    vasRRes.MaxRows = vasRRes.DataRowCnt
    vasRRes.RowHeight(-1) = 13
    
End Sub

Private Sub vasRID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim iCol As Long
    
    iRow = vasRID.ActiveRow
    iCol = vasRID.ActiveCol
    
    If KeyCode = 13 Then
        
        SQL = "update pat_res set pname = '" & Trim(GetText(vasRID, iRow, colPName)) & "' " & vbCrLf & _
              "where equipno = '" & gEquip & "' and isnull(examdate, '') = '" & Format(dtpExamDate, "yyyymmdd") & "' " & vbCrLf & _
              "and isnull(receno, '') = '" & Trim(GetText(vasRID, iRow, colReceno)) & "' and isnull(recedate, '') = '" & Trim(GetText(vasRID, iRow, colReceDate)) & "' " & vbCrLf & _
              "and hospital = '" & Trim(cmbHosp.Text) & "' and sampleno = '" & Trim(GetText(vasRID, iRow, colSampleNo)) & "'"
        res = SendQuery(gLocal, SQL)
    End If
    
    
          
End Sub

Private Sub vasRID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
  PopupMenu Mnlist3
End Sub

Private Sub vasRRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Response, Help
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
        
    vasResRow = vasRRes.ActiveRow
    vasResCol = vasRRes.ActiveCol
    If KeyCode = vbKeyReturn Then
        vasIDRow = vasRID.ActiveRow
        If vasResCol = colResult Then
        
                SQL = " Update pat_res " & vbCrLf & _
                      " Set result = '" & Trim(GetText(vasRRes, vasResRow, colResult)) & "' " & vbCrLf & _
                      " WHERE examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
                      "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                      "  AND equipcode = '" & Trim(GetText(vasRRes, vasResRow, colEquipCode)) & "'" & vbCrLf & _
                      "  AND sampleno = '" & Trim(GetText(vasRID, vasIDRow, colSampleNo)) & "'" & vbCrLf & _
                      "  AND receno = '" & Trim(GetText(vasRID, vasIDRow, colReceno)) & "' "
                res = SendQuery(gLocal, SQL)
                
                SetText vasRRes, Trim(GetText(vasRRes, vasResRow, colResult)), vasResRow, colResult
                
        End If
        
    End If
End Sub

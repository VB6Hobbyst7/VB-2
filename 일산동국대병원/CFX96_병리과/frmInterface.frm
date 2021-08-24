VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   0  '없음
   Caption         =   " CFX96 Interface "
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
   LockControls    =   -1  'True
   Picture         =   "frmInterface.frx":030A
   ScaleHeight     =   10680
   ScaleWidth      =   15165
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   5055
      Left            =   3360
      TabIndex        =   29
      Top             =   4140
      Visible         =   0   'False
      Width           =   14805
      Begin FPSpread.vaSpread vasExcel 
         Height          =   4455
         Left            =   120
         TabIndex        =   60
         Top             =   300
         Width           =   14295
         _Version        =   393216
         _ExtentX        =   25215
         _ExtentY        =   7858
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
      Begin FPSpread.vaSpread vasTemp 
         Height          =   3195
         Left            =   240
         TabIndex        =   39
         Top             =   330
         Width           =   11985
         _Version        =   393216
         _ExtentX        =   21140
         _ExtentY        =   5636
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
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   7230
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
         SpreadDesigner  =   "frmInterface.frx":09BD
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
            Top             =   330
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
         SpreadDesigner  =   "frmInterface.frx":0BD5
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   3210
         TabIndex        =   37
         Top             =   210
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
         SpreadDesigner  =   "frmInterface.frx":0DED
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
         SpreadDesigner  =   "frmInterface.frx":1005
      End
      Begin FPSpread.vaSpread vasResult 
         Height          =   705
         Left            =   1770
         TabIndex        =   50
         Top             =   1860
         Width           =   2175
         _Version        =   393216
         _ExtentX        =   3836
         _ExtentY        =   1244
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
         MaxCols         =   14
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":121D
         UserResize      =   2
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
         SpreadDesigner  =   "frmInterface.frx":1CE7
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
         SpreadDesigner  =   "frmInterface.frx":3760
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
      TabPicture(0)   =   "frmInterface.frx":3978
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":3994
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   8775
         Left            =   180
         TabIndex        =   7
         Top             =   420
         Width           =   14625
         Begin VB.OptionButton optVer 
            Caption         =   "2016"
            Height          =   345
            Index           =   1
            Left            =   4380
            TabIndex        =   62
            Top             =   330
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optVer 
            Caption         =   "2015"
            Height          =   345
            Index           =   0
            Left            =   3510
            TabIndex        =   61
            Top             =   330
            Width           =   855
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   7785
            Left            =   180
            TabIndex        =   14
            Top             =   750
            Width           =   8235
            _Version        =   393216
            _ExtentX        =   14526
            _ExtentY        =   13732
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
            SpreadDesigner  =   "frmInterface.frx":39B0
            UserResize      =   2
         End
         Begin VB.CommandButton cmdExcelFind 
            Caption         =   "엑셀문서 불러오기"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   59
            Top             =   300
            Width           =   2115
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Command1"
            Height          =   315
            Left            =   10350
            TabIndex        =   58
            Top             =   0
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.CommandButton cmdWorkList 
            Caption         =   "WorkList 조회"
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
            Left            =   4470
            TabIndex        =   54
            Top             =   -150
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CommandButton cmdPatDelete 
            Caption         =   "환자 Delete"
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
            Left            =   5880
            TabIndex        =   52
            Top             =   -150
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CommandButton cmdOrder 
            Caption         =   "오더전송"
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
            Left            =   7230
            TabIndex        =   51
            Top             =   -150
            Visible         =   0   'False
            Width           =   1395
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
            SpreadDesigner  =   "frmInterface.frx":443D
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
         Begin MSComCtl2.DTPicker dtpFrDt 
            Height          =   315
            Left            =   1170
            TabIndex        =   53
            Top             =   -90
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   21364737
            CurrentDate     =   40739
         End
         Begin MSComCtl2.DTPicker dtpToDt 
            Height          =   315
            Left            =   2880
            TabIndex        =   55
            Top             =   -90
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   21364737
            CurrentDate     =   40739
         End
         Begin VB.Label Label6 
            Caption         =   "엑셀형식"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   2580
            TabIndex        =   63
            Top             =   420
            Width           =   825
         End
         Begin VB.Label Label3 
            Caption         =   "조회기간 : "
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   57
            Top             =   -30
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label5 
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2700
            TabIndex        =   56
            Top             =   -30
            Visible         =   0   'False
            Width           =   195
         End
      End
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
            Top             =   240
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
            Left            =   3060
            TabIndex        =   24
            Top             =   240
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   240
            TabIndex        =   23
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
            Left            =   780
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
            Left            =   5460
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
            Left            =   6900
            TabIndex        =   18
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   7815
            Left            =   240
            TabIndex        =   21
            Top             =   720
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
            SpreadDesigner  =   "frmInterface.frx":8188
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
            SpreadDesigner  =   "frmInterface.frx":8BBC
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
            TextSave        =   "2016-01-12"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오전 11:40"
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
      Caption         =   "     CFX96 INTERFACE"
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
         Picture         =   "frmInterface.frx":C957
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
Const colA1c = 13
Const colIFCC = 15
Const coleAg = 17

Const calValue = 5.82

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

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Dim OFName As OPENFILENAME

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long


'Dim mOrder.NoOrder  As Boolean
'Dim mOrder.Order    As String
'Dim mOrder.IsSending As Boolean

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
            
            '"GROUP BY BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, JUMIN, Hospital, SENDFLAG"
            
'            SetText vasprint, Trim(GetText(vasrid, iRow, vasrid.MaxCols)), j, 8
'            SetText vasprint, Trim(GetText(vasrid, iRow, 10)), j, 9
            
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
    
    SetForeColor vasID, 1, vasID.maxrows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.maxrows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.maxrows = 0
    vasRes.maxrows = 0
    
'    dtptoday = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            'If Mid(Trim(GetText(vasID, lRow, 3)), 1, 2) = "99" Then
            '    res = Insert_Data_QC(lRow)
            'Else
                res = Insert_Data(lRow)
            'End If
        
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

Private Sub cmdOrder_Click()
    
    intPhase = 1
    strState = "Q"
    intSndPhase = 1
    MSComm1.Output = ENQ
    Save_Raw_Data "[Tx]" & ENQ

End Sub

Private Sub cmdPatDelete_Click()
    
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With vasID
        For i = .DataRowCnt To 1 Step -1
            .Row = i
            .Col = 1
            If .Value = "1" Then
                .Action = ActionDeleteRow
                .maxrows = .maxrows - 1
                j = j + 1
            End If
        Next
    End With

End Sub

Private Sub cmdRClear_Click()
    Dim i As Integer

'    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
    SetForeColor vasRID, 1, vasRID.maxrows, 1, vasRID.MaxCols, 0, 0, 0
    SetForeColor vasRRes, 1, vasRRes.maxrows, 1, vasRRes.MaxCols, 0, 0, 0
    
    vasRID.maxrows = 0
    vasRRes.maxrows = 0
    
    dtpExamDate = Date
    
End Sub

Private Sub cmdRSch_Click()
    Dim iRow As Long

    ClearSpread vasRID
    ClearSpread vasRRes
    Call chkRAll_Click
    
    SQL = "SELECT '', BARCODE, '',RECENO,'', PID, PNAME, PSEX, PAGE " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "GROUP BY BARCODE, RECENO, DISKNO, POSNO, PID, PNAME, PSEX, PAGE"
    res = db_select_Vas(gLocal, SQL, vasRID)
    
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

Private Sub cmdWorkList_Click()
            
    Call GetWorkList(dtpFrDt.Value, dtpToDt.Value)
    vasID.RowHeight(-1) = 12

End Sub




Private Sub cmdSearch_Click()

    Dim idates1$, idates2$, iexamcode$
    Dim pt_no$(), patname$(), sex$(), age$()
    Dim spc_no$(), gnl_item_cd$(), bl_gth_dte$()
    Dim dept$(), wd_no$(), tst_cd$()
    Dim rv As Integer

    Dim ispcno$
    
    Dim liOrdNo As Integer
    Dim i, j, k, n
    Dim lRow, lCol As Long
    Dim lRow1 As Long
    
    Dim sEquipCode As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim lsWorkStation As String
    
    Dim iOrd As Integer
    
    ClearSpread vasID
    
'    If Trim(cboCha.Text) = "" Then
'        MsgBox "차수를 입력하세요!", vbInformation + vbOKOnly, "알림"
'        cboCha.SetFocus
'        Exit Sub
'    End If
    
    Me.MousePointer = 11
    
'    If Not IsDate(dtpSDate.Value) Then
'        dtpSDate.SetFocus
'    End If
'
'    If Not IsDate(dtpEDate.Value) Then
'        dtpEDate.SetFocus
'    End If
    
    sEquipCode = ""
'    IsolateCode cboExam.Text
'    If cboExam.Text = "검사선택" Then
'        sEquipCode = ""
'    Else
'        sEquipCode = cboExam.Text
'    End If
    
'    sEquipCode = Trim(gCode)
    
    If sEquipCode <> "" Then
        sExamCode = ""
        
        SQL = " Select ExamCode From Equipexam Where equipno = '" & gEquip & "' " & CR & _
              " And equipcode = '" & Trim(sEquipCode) & "' "
        res = db_select_Col(gLocal, SQL)
        If res > 0 Then
            For i = 1 To res
                If sExamCode = "" Then
                    sExamCode = "'" & Trim(gReadBuf(0)) & "'"
                Else
                    sExamCode = sExamCode & ", '" & Trim(gReadBuf(0)) & "'"
                End If
            Next i
        End If
    End If
    
    lRow = 1
    lRow1 = 1
    
    idates1$ = Format(dtpFrDt.Value, "yyyymmdd")
    idates2$ = Format(dtpToDt.Value, "yyyymmdd")
    If sExamCode = "" Then
        iexamcode$ = "(" & gAllExam & ")"
    Else
        iexamcode$ = "(" & sExamCode & ")"
    End If
                                  
    rv = sl_d_60_sel_examcode&(idates1$, idates2$, iexamcode$, pt_no$(), patname$(), sex$(), age$(), _
                               spc_no$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
                               
    'ispcno$ = "12020152330" 'spc_no(0)
    'ispcno$ , pt_no$(), patname$(), Sex$(), Age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$()
    'rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), Sex$(), Age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
    If rv < 1 Then
    Else
        vasID.maxrows = rv
        For lRow = 1 To rv
'            sExamName = Get_ExamName(tst_cd(lRow - 1))
'            sEquipCode = Get_EquipCode(tst_cd(lRow - 1))
            SetText vasID, spc_no(lRow - 1), lRow, 3          '검체번호
            SetText vasID, gnl_item_cd$(lRow - 1), lRow, 2    '접수번호
            SetText vasID, pt_no(lRow - 1), lRow, 4           '병록번호
            SetText vasID, patname(lRow - 1), lRow, 5         '이름
            SetText vasID, sExamName, lRow, 6                 '검사명
            SetText vasID, sex(lRow - 1), lRow, 11            '성별
            SetText vasID, age(lRow - 1), lRow, 12            '나이
            SetText vasID, tst_cd(lRow - 1), lRow, 14         '검사코드
            SetText vasID, bl_gth_dte(lRow - 1), lRow, 15         '채취일자
            
            lRow1 = lRow1 + 1
            
'            lRow1 = 1
            
            SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & spc_no(lRow - 1) & "' AND EXAMCODE = '" & tst_cd(lRow - 1) & "' "
            res = db_select_Col(gLocal, SQL)
            
            If res = 0 Then
                SQL = "INSERT INTO PAT_RES(EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE) " & vbCrLf & _
                      "VALUES('" & gEquip & "', '" & spc_no(lRow - 1) & "', '" & gnl_item_cd(lRow - 1) & "', " & vbCrLf & _
                      "'" & pt_no(lRow - 1) & "', '" & patname(lRow - 1) & "', '" & tst_cd(lRow - 1) & "', " & vbCrLf & _
                      "'" & sExamName & "', '" & sex(lRow - 1) & "', '" & age(lRow - 1) & "', '" & Format(Date, "YYYYMMDD") & "', '" & sEquipCode & "')"
                res = SendQuery(gLocal, SQL)
            End If
        
        Next lRow
    End If
    
    '체크,순번,바코드번호,환자번호,환자성명,검사항목,RingID, ,검체,접수,성별,나이,주민번호
    If vasID.maxrows < lRow Then
        vasID.maxrows = lRow
    End If
    
    For lRow = 1 To vasID.DataRowCnt
        iOrd = -1

        SQL = "select barcode, OrdFlag from worklist where barcode = '" & Trim(GetText(vasID, lRow, 3)) & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(1)) = "B" Then
            iOrd = 1
        End If

        
        SQL = "select barcode from pat_res where barcode = '" & Trim(GetText(vasID, lRow, 3)) & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(GetText(vasID, lRow, 3)) Then
            iOrd = 1
        End If

    Next lRow
    
    vasID.maxrows = vasID.DataRowCnt
    
    Me.MousePointer = 0
    
    vasID.RowHeight(-1) = 12
'    SQL = "select '', receno, barcode, pid, pname, examname, '','','','','','','',examcode from pat_Res"
'    res = db_select_Vas(gLocal, SQL, vasid)
    
    
    Exit Sub
End Sub


Private Sub GetWorkList(ByVal pFrDt As String, ByVal pToDt As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strExamCode As String
    Dim j As Integer
    Dim rs As ADODB.Recordset
    Dim sSpecNo As String
    Dim buff As String
    
    buff = "0.7"
    
    vasID.maxrows = 0
    
    '-- 로컬 검사코드 찾기
          SQL = "Select distinct examcode "
    SQL = SQL & "  From EquipExam "
    SQL = SQL & " Where equipno  = '" & Trim(gEquip) & "' "
    
    res = db_select_Row(gLocal, SQL)
    strExamCode = ""
    For i = 0 To UBound(gReadBuf)
        If gReadBuf(i) <> "" Then
            strExamCode = strExamCode & "'" & Trim(gReadBuf(i)) & "',"
        Else
            Exit For
        End If
    Next
    strExamCode = Mid(strExamCode, 1, Len(strExamCode) - 1)
    
    
    
    
    
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    '-- 결과검증
'    Dim oerrmsg As String
'    Dim ispcid(1 To 1)  As String
'    Dim iexamcode(1 To 1) As String, iresult(1 To 1) As String
'    Dim ierrflag(1 To 1) As String, iequipcd(1 To 1) As String
'    Dim iuser(1 To 1) As String
    
    
    Dim idates1(1 To 1) As String
    Dim idates2(1 To 1) As String
    Dim iexamcode(1 To 1) As String
'    Dim strpt_no(1 To 1) As String
'    Dim strpt_nm(1 To 1) As String
'    Dim strSex(1 To 1) As String
'    Dim strage(1 To 1) As String
'    Dim strspc_no(1 To 1) As String
'    Dim strgnl_item_cd(1 To 1) As String
'    Dim strbl_gth_dte(1 To 1) As String
'    Dim strdept(1 To 1) As String
'    Dim strwd_no(1 To 1) As String
'    Dim strtst_cd(1 To 1) As String
    Dim strpt_no() As String
    Dim strpt_nm() As String
    Dim strSex() As String
    Dim strAge() As String
    Dim strspc_no() As String
    Dim strgnl_item_cd() As String
    Dim strbl_gth_dte() As String
    Dim strdept() As String
    Dim strwd_no() As String
    Dim strtst_cd() As String
    
'    ispcid(1) = strSpcid
'    iexamcode(1) = strExamCode
'    iresult(1) = strResult
'    ierrflag(1) = "0"
'    iequipcd(1) = strEquipcd
'    iuser$(1) = strUser
    
    '-- old
    'oerrmsg$ , ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$
    'Call sl_online_result_ul_4&(oerrmsg, ispcid, iexamcode, iresult, ierrflag, iequipcd, igubun)
idates1(1) = "20120101"
idates2(1) = "20120203"
iexamcode(1) = "L905261"
'strpt_no(1) = ""
'strpt_nm(1) = ""

Dim varAry As Variant
Dim strAry As String
    '-- new
    'oerrmsg$, ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), iuser$()
    varAry = sl_d_60_sel_examcode&("20120101", "20120203", "L905273", strpt_no, strpt_nm, strSex, strAge, strspc_no, strgnl_item_cd, strbl_gth_dte, strdept, strwd_no, strtst_cd)
    
    strAry = varAry
'select  a.pt_no           pt_no,
'        c.patname         patname,
'      c.sex            sex,
'        f_age(c.birthday) age,
'        a.spc_no          spc_no,
'        a.gnl_item_cd     gnl_item_cd,
'        to_date(d.acpt_dte, 'yyyy-mm-dd')      bl_gth_dte,
'        d.dept            dept,
'        d.wd_no           wd_no,
'        a.tst_cd tst_cd
    
    
    '-- 검사대상자 가져오기
    SQL = "Select distinct SPCM_NO,WORK_NO  From SPSLHRRST " & CR & _
          " Where rcpn_dt between '" & Format(pFrDt, "yyyymmdd") & "' and '" & Format(pToDt, "yyyymmdd") & "'" & vbCrLf & _
          "   and exmn_cd in (" & strExamCode & ")" & _
          "   and rslt_no IS NOT NULL" & _
          " Order by WORK_NO,SPCM_NO"
  
    Set rs = cn_Ser.Execute(SQL, , 1)
          
    Do Until rs.EOF
        SQL = "SELECT FN_LABCVTPRTBCNO('" & Trim(rs.Fields(0)) & "') FROM DUAL "
        res = db_select_Col(gServer, SQL)
        sSpecNo = Trim(gReadBuf(0))
        
        SQL = "SELECT PID, PT_NM, SEX, AGE "
        SQL = SQL & vbCrLf & " FROM SPSLMJBBI "
        SQL = SQL & vbCrLf & "WHERE SPCM_NO = '" & Trim(rs.Fields(0)) & "' "
        SQL = SQL & vbCrLf & "  AND SPCM_STAT = '2' "
        SQL = SQL & vbCrLf & "  AND RSLT_STAT <= '1' "
        res = db_select_Col(gServer, SQL)
        
        If res > 0 Then
            j = j + 1
            vasID.maxrows = j
            SetText vasID, 1, j, 1     '2
            SetText vasID, Trim(rs.Fields(0)), j, colSpecNo     '2
            SetText vasID, sSpecNo, j, colBarcode     '3
            SetText vasID, Format(Mid(Trim(rs.Fields(0)), 3, 6), "######"), j, colRack  '4
            SetText vasID, Format(j, "0000"), j, colPos    '5
            SetText vasID, Trim(gReadBuf(0)), j, colPID    '6
            SetText vasID, Trim(gReadBuf(1)), j, colPName  '7
            SetText vasID, Trim(gReadBuf(2)), j, colSex    '8
            SetText vasID, Trim(gReadBuf(3)), j, colAge    '9
        End If
        rs.MoveNext
    
    Loop
    
    vasID.RowHeight(-1) = 12

End Sub

Public Function Get_ExamCode(asExamName As String) As String
    Dim strExamName As String
    Get_ExamCode = ""
    strExamName = Trim(asExamName)
        
    SQL = "SELECT EXAMCODE FROM EQUIPEXAM WHERE EQUIPCODE = '" & asExamName & "'"
    res = db_select_Col(gLocal, SQL)
    
    If res < 0 Then
        SaveQuery SQL
        Exit Function
    End If
    
    Get_ExamCode = Trim(gReadBuf(0))
    
End Function

Public Function Get_ExamName(asExamCode As String) As String
    Dim strExamCode As String
    Get_ExamName = ""
    strExamCode = Trim(asExamCode)
        
    SQL = "SELECT EXAMNAME FROM EQUIPEXAM WHERE EXAMCODE = '" & strExamCode & "'"
    res = db_select_Col(gLocal, SQL)
    
    If res < 0 Then
        SaveQuery SQL
        Exit Function
    End If
    
    Get_ExamName = Trim(gReadBuf(0))
    
End Function


Public Function Get_EquipCode(asExamCode As String) As String
    Dim strExamCode As String
    Get_EquipCode = ""
    strExamCode = Trim(asExamCode)
        
    SQL = "SELECT EQUIPCODE FROM EQUIPEXAM WHERE EXAMCODE = '" & strExamCode & "'"
    res = db_select_Col(gLocal, SQL)
    
    If res < 0 Then
        SaveQuery SQL
        Exit Function
    End If
    
    Get_EquipCode = Trim(gReadBuf(0))
    
End Function


Private Function ShowOpen(Ufilter As String, Upath As String) As String
    
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hWnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Ufilter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = Upath
    OFName.lpstrTitle = "Open File"
    OFName.flags = 0

    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
        'ShowOpen = Mid(ShowOpen, 1, Len(ShowOpen) - 1)
    Else
        ShowOpen = ""
    End If
    
End Function

Private Sub getCFXExlData(ByVal intIDX As Integer)

    Dim x As Integer, y As Boolean, z As Boolean
    Dim ListCount   As Integer, handle As Integer
    Dim List(10)    As String
    Dim intRow, intCol As Long
    Dim varTmp      As Variant
    Dim strExcel    As String
    Dim rv          As Integer
    Dim lRow        As Integer
    Dim lRow1       As Integer
    Dim sExamCode   As String
    Dim sExamName   As String
    Dim sEquipCode  As String
    Dim sItemCode   As String
    Dim strAge      As String
    Dim strSex      As String
    Dim strPtno     As String
    Dim strPtname   As String
    Dim strTmp      As String
    Dim sFile       As String
'    Dim intSheet    As Integer

    Dim idates1$, idates2$, iexamcode$
    Dim pt_no$(), patname$(), sex$(), age$()
    Dim spc_no$(), gnl_item_cd$(), bl_gth_dte$()
    Dim dept$(), wd_no$(), tst_cd$()
    Dim ispcno$

'On Error Resume Next

    '-- 이전버전
    If intIDX = 0 Then
        sFile = ShowOpen("Excel Files (*.xls)|*.xls|All Files (*.*)|*.*", App.Path)
        If sFile <> "" Then
            vasID.maxrows = 0
            strExcel = sFile
            vasExcel.ScriptEnhanced = True
            x = vasExcel.IsExcelFile(strExcel)
            If x = 1 Then
                y = vasExcel.GetExcelSheetList(strExcel, List, ListCount, "Report.txt", handle, True)
                If y = True Then
                    z = vasExcel.ImportExcelSheet(handle, 0)
                    If z = True Then
                        'MsgBox "가져오기 성공"
                    Else
                        'MsgBox "가려오기 실패"
                    End If
                End If
            End If
            
            With vasExcel
                For intRow = 2 To .DataRowCnt
                    For intCol = 1 To 5
                        .GetText intCol, intRow, varTmp
                        If varTmp <> "" Then
                            Select Case intCol
                            Case 1
                                ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
                                rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), sex$(), age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
                                'rv = 1
                                If rv >= 1 Then
                                    '-- 환자정보
                                    vasID.maxrows = vasID.maxrows + 1
                                    vasID.RowHeight(-1) = 12
                                    lRow = vasID.maxrows
                                    sExamName = Get_ExamName(tst_cd(0))
                                    sEquipCode = Get_EquipCode(tst_cd(0))
                                    sItemCode = gnl_item_cd(0)
                                    strPtno = patname(0)
                                    strPtname = patname(0)
                                    strSex = sex(0)
                                    strAge = age(0)
                                    
                                    SetText vasID, ispcno$, lRow, 2
                                    SetText vasID, gnl_item_cd(0), lRow, 4         '채취일자
                                    SetText vasID, pt_no(0), lRow, 6           '병록번호
                                    SetText vasID, patname(0), lRow, 7         '이름
                                    SetText vasID, sex(0), lRow, 8            '성별
                                    SetText vasID, age(0), lRow, 9            '나이
                                    '-- 채널
                                    .GetText 5, intRow, varTmp: sEquipCode = varTmp
                                    sExamCode = Get_ExamCode(sEquipCode)
                                    sExamName = Get_ExamName(sExamCode)
                                    
                                          SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
                                    SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
                                    res = db_select_Col(gLocal, SQL)
                                    '-- 결과
                                    .GetText 6, intRow, varTmp: strTmp = varTmp
                                    If strTmp = "-" Then
                                        strTmp = "Negative"
                                    ElseIf strTmp = "+" Then
                                        strTmp = "Positive(1+)"
                                    ElseIf strTmp = "++" Then
                                        strTmp = "Positive(2+)"
                                    ElseIf strTmp = "+++" Then
                                        strTmp = "Positive(3+)"
                                    ElseIf strTmp = "++++" Then
                                        strTmp = "Positive(4+)"
                                    'Else
                                    '    strTmp = "Positive"
                                    End If
                                    
                                    If res > 0 Then
                                        SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                                              "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
                                              "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                                              "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
                                              "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
                                              "  and examcode= '" & sExamCode & "'"
                                        res = SendQuery(gLocal, SQL)
                                    End If
                                    
                                    SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
                                          "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
                                          "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
                                          "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
                                    res = SendQuery(gLocal, SQL)
    '                                Debug.Print SQL
                                    
                                    Exit For
                                End If
                            Case "5"
                                    '-- 채널
                                    .GetText 5, intRow, varTmp: sEquipCode = varTmp
                                    sExamCode = Get_ExamCode(sEquipCode)
                                    sExamName = Get_ExamName(sExamCode)
    
                                          SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
                                    SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
                                    res = db_select_Col(gLocal, SQL)
                                    '-- 결과
                                    .GetText 6, intRow, varTmp: strTmp = varTmp
                                    If strTmp = "-" Then
                                        strTmp = "Negative"
                                    ElseIf strTmp = "+" Then
                                        strTmp = "Positive(1+)"
                                    ElseIf strTmp = "++" Then
                                        strTmp = "Positive(2+)"
                                    ElseIf strTmp = "+++" Then
                                        strTmp = "Positive(3+)"
                                    ElseIf strTmp = "++++" Then
                                        strTmp = "Positive(4+)"
                                    'Else
                                    '    strTmp = "Positive"
                                    End If
                                    
                                    
                                    If res > 0 Then
                                        SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                                              "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
                                              "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                                              "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
                                              "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
                                              "  and examcode= '" & sExamCode & "'"
                                        res = SendQuery(gLocal, SQL)
                                    End If
                                    
                                    SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
                                          "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
                                          "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
                                          "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
                                    res = SendQuery(gLocal, SQL)
    '                                Debug.Print SQL
                                    Exit For
                            End Select
                        End If
                    Next
                Next
            End With
        End If
    
    Else
    '-- 신버전
        sFile = ShowOpen("Excel Files (*.xls)|*.xls|All Files (*.*)|*.*", App.Path)
        If sFile <> "" Then
            vasID.maxrows = 0
            strExcel = sFile
            vasExcel.ScriptEnhanced = True
            x = vasExcel.IsExcelFile(strExcel)
            If x = 1 Then
                y = vasExcel.GetExcelSheetList(strExcel, List, ListCount, "Report.txt", handle, True)
                If y = True Then
                    '-- 2번째sheet
                    z = vasExcel.ImportExcelSheet(handle, 1)
                    If z = True Then
                        'MsgBox "가져오기 성공"
                        With vasExcel
                            For intRow = 2 To .DataRowCnt Step 4
                                .GetText 1, intRow, varTmp
                                ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
                                rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), sex$(), age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
                                'rv = 1
                                If rv >= 1 Then
                                    '-- 환자정보
                                    vasID.maxrows = vasID.maxrows + 1
                                    vasID.RowHeight(-1) = 12
                                    lRow = vasID.maxrows
                                    sExamName = Get_ExamName(tst_cd(0))
                                    sEquipCode = Get_EquipCode(tst_cd(0))
                                    sItemCode = gnl_item_cd(0)
                                    strPtno = patname(0)
                                    strPtname = patname(0)
                                    strSex = sex(0)
                                    strAge = age(0)
                                    
                                    SetText vasID, ispcno$, lRow, 2
                                    SetText vasID, gnl_item_cd(0), lRow, 4         '채취일자
                                    SetText vasID, pt_no(0), lRow, 6           '병록번호
                                    SetText vasID, patname(0), lRow, 7         '이름
                                    SetText vasID, sex(0), lRow, 8            '성별
                                    SetText vasID, age(0), lRow, 9            '나이
                                End If
                            
                                For intCol = 6 To 22 Step 2
                                    .GetText intCol, intRow, varTmp
                                    If varTmp <> "" Then
                                        '-- Row = 1,2 채널
                                        .GetText intCol, intRow, varTmp: sEquipCode = varTmp
                                        sExamCode = Get_ExamCode(sEquipCode)
                                        sExamName = Get_ExamName(sExamCode)
                
                                              SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
                                        SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
                                        res = db_select_Col(gLocal, SQL)
                                        '-- 결과
                                        .GetText intCol, intRow + 1, varTmp: strTmp = varTmp
                                        If strTmp = "-" Then
                                            strTmp = "Negative"
                                        ElseIf strTmp = "+" Then
                                            strTmp = "Positive(1+)"
                                        ElseIf strTmp = "++" Then
                                            strTmp = "Positive(2+)"
                                        ElseIf strTmp = "+++" Then
                                            strTmp = "Positive(3+)"
                                        ElseIf strTmp = "++++" Then
                                            strTmp = "Positive(4+)"
                                        'Else
                                        '    strTmp = "Positive"
                                        End If
                                        
                                        If res > 0 Then
                                            SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                                                  "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
                                                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                                                  "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
                                                  "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
                                                  "  and examcode= '" & sExamCode & "'"
                                            res = SendQuery(gLocal, SQL)
                                        End If
                                        
                                        SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
                                              "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
                                              "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
                                              "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
                                        res = SendQuery(gLocal, SQL)
                                        
                                        '-- Row = 3,4 채널
                                        .GetText intCol, intRow + 2, varTmp: sEquipCode = varTmp
                                        sExamCode = Get_ExamCode(sEquipCode)
                                        sExamName = Get_ExamName(sExamCode)
                
                                              SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
                                        SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
                                        res = db_select_Col(gLocal, SQL)
                                        '-- 결과
                                        .GetText intCol, intRow + 3, varTmp: strTmp = varTmp
                                        If strTmp = "-" Then
                                            strTmp = "Negative"
                                        ElseIf strTmp = "+" Then
                                            strTmp = "Positive(1+)"
                                        ElseIf strTmp = "++" Then
                                            strTmp = "Positive(2+)"
                                        ElseIf strTmp = "+++" Then
                                            strTmp = "Positive(3+)"
                                        ElseIf strTmp = "++++" Then
                                            strTmp = "Positive(4+)"
                                        'Else
                                        '    strTmp = "Positive"
                                        End If
                                        
                                        If res > 0 Then
                                            SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                                                  "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
                                                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                                                  "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
                                                  "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
                                                  "  and examcode= '" & sExamCode & "'"
                                            res = SendQuery(gLocal, SQL)
                                        End If
                                        
                                        SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
                                              "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
                                              "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
                                              "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
                                        res = SendQuery(gLocal, SQL)
                                    End If
                                Next
                            Next
                        End With
                    End If
                    '-- 3번째sheet
                    z = vasExcel.ImportExcelSheet(handle, 2)
                    If z = True Then
                        'MsgBox "가져오기 성공"
                        With vasExcel
                            For intRow = 3 To .DataRowCnt
                                .GetText 1, intRow, varTmp
                                ispcno$ = varTmp '"12020356983" '"12020152330" ' "12020356983" ''varTmp
                                rv = sl_d_60_sel_spcno&(ispcno$, pt_no$(), patname$(), sex$(), age$(), gnl_item_cd$(), bl_gth_dte$(), dept$(), wd_no$(), tst_cd$())
                                'rv = 1
                                If rv >= 1 Then
                                    '-- 환자정보
                                    vasID.maxrows = vasID.maxrows + 1
                                    vasID.RowHeight(-1) = 12
                                    lRow = vasID.maxrows
                                    sExamName = Get_ExamName(tst_cd(0))
                                    sEquipCode = Get_EquipCode(tst_cd(0))
                                    sItemCode = gnl_item_cd(0)
                                    strPtno = patname(0)
                                    strPtname = patname(0)
                                    strSex = sex(0)
                                    strAge = age(0)
                                    
                                    SetText vasID, ispcno$, lRow, 2
                                    SetText vasID, gnl_item_cd(0), lRow, 4         '채취일자
                                    SetText vasID, pt_no(0), lRow, 6           '병록번호
                                    SetText vasID, patname(0), lRow, 7         '이름
                                    SetText vasID, sex(0), lRow, 8            '성별
                                    SetText vasID, age(0), lRow, 9            '나이
                                End If
                            
                                For intCol = 6 To 20 Step 2
                                    .GetText intCol, 2, varTmp
                                    If varTmp <> "" Then
                                        '-- Row = 1,2 채널
                                        .GetText intCol, 2, varTmp: sEquipCode = varTmp
                                        sExamCode = Get_ExamCode(sEquipCode)
                                        sExamName = Get_ExamName(sExamCode)
                
                                              SQL = "SELECT EXAMCODE FROM PAT_RES WHERE BARCODE = '" & ispcno$ & "' "
                                        SQL = SQL & "   AND EXAMCODE = '" & sExamCode & "' "
                                        res = db_select_Col(gLocal, SQL)
                                        '-- 결과
                                        .GetText intCol, intRow, varTmp: strTmp = varTmp
                                        If strTmp = "-" Then
                                            strTmp = "Negative"
                                        ElseIf strTmp = "+" Then
                                            strTmp = "Positive(1+)"
                                        ElseIf strTmp = "++" Then
                                            strTmp = "Positive(2+)"
                                        ElseIf strTmp = "+++" Then
                                            strTmp = "Positive(3+)"
                                        ElseIf strTmp = "++++" Then
                                            strTmp = "Positive(4+)"
                                        'Else
                                        '    strTmp = "Positive"
                                        End If
                                        
                                        If res > 0 Then
                                            SQL = "DELETE FROM PAT_RES " & vbCrLf & _
                                                  "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
                                                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                                                  "  AND BARCODE = '" & ispcno$ & "' " & vbCrLf & _
                                                  "  and equipcode = '" & sEquipCode & "'" & vbCrLf & _
                                                  "  and examcode= '" & sExamCode & "'"
                                            res = SendQuery(gLocal, SQL)
                                        End If
                                        
                                        SQL = "INSERT INTO PAT_RES( EQUIPNO, BARCODE, RECENO, PID, PNAME, EXAMCODE, EXAMNAME, PSEX, PAGE, EXAMDATE, EQUIPCODE,RESULT) " & vbCrLf & _
                                              "VALUES('" & gEquip & "', '" & ispcno$ & "', '" & sItemCode & "', " & vbCrLf & _
                                              "'" & strPtno & "', '" & strPtname & "', '" & sExamCode & "', " & vbCrLf & _
                                              "'" & sExamName & "', '" & strSex & "', '" & strAge & "', '" & Format(dtpToday, "YYYYMMDD") & "', '" & sEquipCode & "','" & strTmp & "')"
                                        res = SendQuery(gLocal, SQL)
                                        
                                        
                                    End If
                                Next
                            Next
                        End With
                    End If
                    
                End If
            End If
        End If
    End If

End Sub

Private Sub cmdExcelFind_Click()

    If optVer(0).Value = True Then
        Call getCFXExlData(0)
    Else
        Call getCFXExlData(1)
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
    
    
            strBuffer = ""
strBuffer = strBuffer & "1H|\^&||||||||||P||" & vbCrLf
strBuffer = strBuffer & "05" & vbCrLf
strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||||||||||" & vbCrLf
strBuffer = strBuffer & "3B" & vbCrLf
strBuffer = strBuffer & "3O|1|1551500020|113^0005^3^^SAMPLE^NORMAL|ALL|R|20041029092240|||||X||||||||||||||O|||||" & vbCrLf
strBuffer = strBuffer & "A3" & vbCrLf
strBuffer = strBuffer & "4R|1|^^^440^^0|-1^0.324|COI|^|N||F|||20041029101118|20041029102940|" & vbCrLf
strBuffer = strBuffer & "AD" & vbCrLf
strBuffer = strBuffer & "5R|2|^^^430^^0|>1000|IU/l|2.00^10.00|>||F|||20041029101200|20041029103022|" & vbCrLf
strBuffer = strBuffer & "E9" & vbCrLf
strBuffer = strBuffer & "6C|1|I|51^Above measuring range|I" & vbCrLf
strBuffer = strBuffer & "05" & vbCrLf
strBuffer = strBuffer & "7R|3|^^^50^^0|1.19|ng/ml|0.800^2.00|N||F|||20041029101242|20041029103104|" & vbCrLf
strBuffer = strBuffer & "48" & vbCrLf
strBuffer = strBuffer & "0R|4|^^^21^^0|8.46|ug/dl|5.10^14.10|N||F|||20041029101324|20041029103146|" & vbCrLf
strBuffer = strBuffer & "4E" & vbCrLf
strBuffer = strBuffer & "1R|5|^^^10^^0|1.14|uIU/ml|0.270^4.20|N||F|||20041029101406|20041029103228|" & vbCrLf
strBuffer = strBuffer & "87" & vbCrLf
strBuffer = strBuffer & "2L|1" & vbCrLf
strBuffer = strBuffer & "3B" & vbCrLf
strBuffer = strBuffer & "" & vbCrLf
    
'    4R|1|^^^HB2CAP96|99.932991027832|IU/mL|20^170000000|N||V||LIS|20110829124243|20110829145822|39-1588
'D0
'
'18:20:49 [Tx]  < 2.00E+1
'4R|1|^^^HB2CAP96|1.21385717391968|IU/mL|20^170000000|<||V||LIS|20110829124243|20110829145822|39-1588
'EC
'
'
'18:22:04 [Tx]  2.03E+2
'4R|1|^^^HB2CAP96|203.288360595703|IU/mL|20^170000000|N||V||LIS|20110829124243|20110829145822|39-1588
'F4

            strBuffer = ENQ & STX & "1H|\^&|||ALCZC6284HXW^Roche^AMPLILINK^3.2.2.0801^Roche ASTM+^CZC6284HXW||||||||1|20110825165130" & vbCr & ETX & "4F" & vbCr & vbLf
strBuffer = strBuffer & STX & "2P|1||10877204||Lee sang jin|||||||||||0^cm|0^kg" & vbCr & ETX & "9A" & vbCr & vbLf
strBuffer = strBuffer & STX & "3O|1|1605900070|110704-0014|^^^ALL|R|20110704154308|||0^ml||A" & vbCr & ETX & "EF" & vbCr & vbLf
strBuffer = strBuffer & STX & "4R|1|^^^HB2CAP96|99.932991027832|||N||X||LIS|20110704162921|20110704162921|39-2579" & vbCr & ETX & "EA" & vbCr & vbLf    '9.99E+1
strBuffer = strBuffer & STX & "4R|1|^^^HB2CAP96|1.21385717391968|||<||X||LIS|20110704162921|20110704162921|39-2579" & vbCr & ETX & "EA" & vbCr & vbLf   '< 2.00E+1
strBuffer = strBuffer & STX & "4R|1|^^^HB2CAP96|203.288360595703|||N||X||LIS|20110704162921|20110704162921|39-2579" & vbCr & ETX & "EA" & vbCr & vbLf  '2.03E+2
strBuffer = strBuffer & STX & "4R|1|^^^HB2CAP96|1304756736|IU/mL|20^170000000|>||V||LIS|20110822153437|20110822175030|39-1588" & vbCr & ETX & "EA" & vbCr & vbLf


strBuffer = strBuffer & STX & "4R|1|^^^HB2CAP96|Target Not Detected||20^170000000|N||V||LIS|20110818135405|20110818161000|39-1588" & vbCr & ETX & "EA" & vbCr & vbLf


'strBuffer = strBuffer & STX & "4R|1|^^^HB2CAP96|> 1.70E+8 IU/mL|||A||X||LIS|20110704162921|20110704162921|39-2579" & vbCr & ETX & "EA" & vbCr & vbLf
'strBuffer = strBuffer & STX & "4R|1|^^^HB2CAP96|1.01E+4 IU/mL|||A||X||LIS|20110704162921|20110704162921|39-2579" & vbCr & ETX & "EA" & vbCr & vbLf
strBuffer = strBuffer & STX & "5C|1|I|AP12^ Sample Input missing|I" & vbCr & ETX & "59" & vbCr & vbLf
strBuffer = strBuffer & STX & "6C|2|I|AP24^ Preparation aborted|I" & vbCr & ETX & "38" & vbCr & vbLf
strBuffer = strBuffer & STX & "7L|1|N" & vbCr & ETX & "OA" & vbCr & vbLf
strBuffer = strBuffer & EOT

'strBuffer = ACK
    
strBuffer = ""
strBuffer = strBuffer & ""
strBuffer = strBuffer & "1H|\^&|||ALCZC6284HXW^Roche^AMPLILINK^3.3.5.1002^Roche ASTM+^CZC6284HXW||||||||1|20111007152334" & vbCr & vbLf
strBuffer = strBuffer & "48" & vbCr & vbLf
strBuffer = strBuffer & "2P|1||10000275||20111006L900001l|||||||||||0^cm|0^kg" & vbCr & vbLf
strBuffer = strBuffer & "B9" & vbCr & vbLf
strBuffer = strBuffer & "3O|1|1120011131|111006-0001|^^^ALL|R|20111007113658|||0^ml||A" & vbCr & vbLf
strBuffer = strBuffer & "CA" & vbCr & vbLf
strBuffer = strBuffer & "4R|1|^^^HB2CAP96|8.42E+1|IU/mL|20^170000000^TiterRanges|N||V||LIS|20111007124923|20111007150506|CTM1" & vbCr & vbLf
strBuffer = strBuffer & "7C" & vbCr & vbLf
strBuffer = strBuffer & "5C|1||Accepted|G" & vbCr & vbLf
strBuffer = strBuffer & "09" & vbCr & vbLf
strBuffer = strBuffer & "6L|1|N" & vbCr & vbLf
strBuffer = strBuffer & "09" & vbCr & vbLf
strBuffer = strBuffer & ""


strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||ALCZC6284HXW^Roche^AMPLILINK^3.3.5.1002^Roche ASTM+^CZC6284HXW||||||||1|20111007152334" & vbCr & vbLf
strBuffer = strBuffer & "48" & vbCr & vbLf
strBuffer = strBuffer & "2P|1||10690391||20111006L900003l|||||||||||0^cm|0^kg" & vbCr & vbLf
strBuffer = strBuffer & "C9" & vbCr & vbLf
strBuffer = strBuffer & "3O|1|1120011191|111006-0003|^^^ALL|R|20111007113700|||0^ml||A" & vbCr & vbLf
strBuffer = strBuffer & "C6" & vbCr & vbLf
strBuffer = strBuffer & "4R|1|^^^HB2CAP96|1.43E+6|IU/mL|20^170000000^TiterRanges|N||V||LIS|20111007124923|20111007150506|CTM1" & vbCr & vbLf
strBuffer = strBuffer & "7B" & vbCr & vbLf
strBuffer = strBuffer & "5C|1||Accepted|G" & vbCr & vbLf
strBuffer = strBuffer & "09" & vbCr & vbLf
strBuffer = strBuffer & "6L|1|N" & vbCr & vbLf
strBuffer = strBuffer & "09" & vbCr & vbLf
strBuffer = strBuffer & ""

strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||ALCZC6284HXW^Roche^AMPLILINK^3.3.5.1002^Roche ASTM+^CZC6284HXW||||||||1|20111007152336" & vbCr & vbLf
strBuffer = strBuffer & "4A" & vbCr & vbLf
strBuffer = strBuffer & "2P|1||10840068||20111006L900002l|||||||||||0^cm|0^kg" & vbCr & vbLf
strBuffer = strBuffer & "C6" & vbCr & vbLf
strBuffer = strBuffer & "3O|1|1120011171|111006-0002|^^^ALL|R|20111007113659|||0^ml||A" & vbCr & vbLf
strBuffer = strBuffer & "D0" & vbCr & vbLf
strBuffer = strBuffer & "4R|1|^^^HB2CAP96|> 1.70E+8|IU/mL|20^170000000^TiterRanges|>||V||LIS|20111007124923|20111007150506|CTM1" & vbCr & vbLf
strBuffer = strBuffer & "CB" & vbCr & vbLf
strBuffer = strBuffer & "5C|1||Accepted|G" & vbCr & vbLf
strBuffer = strBuffer & "09" & vbCr & vbLf
strBuffer = strBuffer & "6C|2|I|TM23^ ABOVE_RANGE|I" & vbCr & vbLf
strBuffer = strBuffer & "FA" & vbCr & vbLf
strBuffer = strBuffer & "7L|1|N" & vbCr & vbLf
strBuffer = strBuffer & "0A" & vbCr & vbLf
strBuffer = strBuffer & ""

strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||ALCZC6284HXW^Roche^AMPLILINK^3.3.5.1002^Roche ASTM+^CZC6284HXW||||||||1|20111007162815" & vbCr & vbLf
strBuffer = strBuffer & "4D" & vbCr & vbLf
strBuffer = strBuffer & "2P|1||10574913||Kim mi ja|||||||||||0^cm|0^kg" & vbCr & vbLf
strBuffer = strBuffer & "5D" & vbCr & vbLf
strBuffer = strBuffer & "3O|1|1118441931|110929-0027|^^^ALL|R|20110929112024|||0^ml||A" & vbCr & vbLf
strBuffer = strBuffer & "F3" & vbCr & vbLf
strBuffer = strBuffer & "4R|1|^^^HB2CAP96|< 2.00E+1|IU/mL|20^170000000^TiterRanges|<||V||LIS|20110929142852|20110929163752|CTM1" & vbCr & vbLf
strBuffer = strBuffer & "DA" & vbCr & vbLf
strBuffer = strBuffer & "5C|1||Accepted|G" & vbCr & vbLf
strBuffer = strBuffer & "09" & vbCr & vbLf
strBuffer = strBuffer & "6C|2|I|TM22^ BELOW_RANGE|I" & vbCr & vbLf
strBuffer = strBuffer & "05" & vbCr & vbLf
strBuffer = strBuffer & "7L|1|N" & vbCr & vbLf
strBuffer = strBuffer & "0A" & vbCr & vbLf
strBuffer = strBuffer & ""
    

'-- iNVALID
strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||ALCZC6284HXW^Roche^AMPLILINK^3.3.5.1002^Roche ASTM+^CZC6284HXW||||||||1|20111014120033" & vbCr & vbLf
strBuffer = strBuffer & "3D" & vbCr & vbLf
strBuffer = strBuffer & "2P|1||10490316||Chang jeong man|||||||||||0^cm|0^kg" & vbCr & vbLf
strBuffer = strBuffer & "C5" & vbCr & vbLf
strBuffer = strBuffer & "3O|1|1118268351|111004-0029|^^^ALL|R|20111004103510|||0^ml||A" & vbCr & vbLf
strBuffer = strBuffer & "DA" & vbCr & vbLf
strBuffer = strBuffer & "4R|1|^^^HB2CAP96|Invalid||20^170000000^TiterRanges|A||X||LIS|20111004113659|20111004135241|CTM1" & vbCr & vbLf
strBuffer = strBuffer & "42" & vbCr & vbLf
strBuffer = strBuffer & "5C|1||Accepted|G" & vbCr & vbLf
strBuffer = strBuffer & "09" & vbCr & vbLf
strBuffer = strBuffer & "6C|2|I|TM45^ DRIFT_HIGH-1|I" & vbCr & vbLf
strBuffer = strBuffer & "1B" & vbCr & vbLf
strBuffer = strBuffer & "7C|3|I|TM42^ SPK_CORR-1|I" & vbCr & vbLf
strBuffer = strBuffer & "A5" & vbCr & vbLf
strBuffer = strBuffer & "0L|1|N" & vbCr & vbLf
strBuffer = strBuffer & "03" & vbCr & vbLf
strBuffer = strBuffer & ""



strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||ALCZC6284HXW^Roche^AMPLILINK^3.3.5.1002^Roche ASTM+^CZC6284HXW||||||||1|2011110411073343" & vbCr & vbLf
strBuffer = strBuffer & "2P|1||10263471||20111102L900003l|||||||||||0^cm|0^kgC1" & vbCr & vbLf
strBuffer = strBuffer & "3O|1|1120048161|111102-0002|^^^ALL|R|20111102180258|||0^ml||AD1" & vbCr & vbLf
strBuffer = strBuffer & "4R|1|^^^HB2CAP96|6.94E+2|IU/mL|20^170000000^TiterRanges|N||V||LIS|20111102185754|20111102211335|CTM181" & vbCr & vbLf
strBuffer = strBuffer & "5C|1||Accepted|G09" & vbCr & vbLf
strBuffer = strBuffer & "6L|1|N09" & vbCr & vbLf
strBuffer = strBuffer & "" & vbCr & vbLf


strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||ALCZC6284HXW^Roche^AMPLILINK^3.3.5.1002^Roche ASTM+^CZC6284HXW||||||||1|2011110715010644" & vbCr & vbLf
strBuffer = strBuffer & "2P|1||10331568||20111105L900012l|||||||||||0^cm|0^kgC7" & vbCr & vbLf
strBuffer = strBuffer & "3O|1|1120037121|111105-0007|^^^ALL|R|20111107110316|||0^ml||ACC" & vbCr & vbLf
strBuffer = strBuffer & "4R|1|^^^HB2CAP96|5.69E+5|IU/mL|20^170000000^TiterRanges|N||V||LIS|20111107122242|20111107143824|CTM185" & vbCr & vbLf
strBuffer = strBuffer & "5C|1||Accepted|G09" & vbCr & vbLf
strBuffer = strBuffer & "6L|1|N09" & vbCr & vbLf
strBuffer = strBuffer & "" & vbCr & vbLf

'Call MSComm1_OnComm


    
    Dim wkbuf As String
    
'    Open App.Path & "\log\long.log" For Input As #3
    Open App.Path & "\log\1110.log" For Input As #3
    
    wkbuf = ""
    
    Do While Not EOF(3)
        wkbuf = wkbuf & Input(1, #3)
    Loop

    Close #3

    strBuffer = wkbuf
    
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
        If llRow > vasID.maxrows Then
            vasID.maxrows = llRow
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

    vasRes.maxrows = Trim(gReadBuf(0))

    
    lsSeq = ""
    lsCnt = ""
        
    
    '결과 잘라 넣기
    j = 0
    For j = 1 To vasRes.maxrows
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




Private Sub E411(asData As String)
    
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim i As Integer
    Dim ii As Integer
    Dim j As Integer
    Dim k As Integer
    Dim x As Integer
    
    Dim iCnt As Integer
    
    Dim lsID As String
    Dim lsPid As String
    Dim lsPName As String
    Dim lsJumin1 As String
    Dim lsJumin2 As String
    Dim lsPSex As String
    Dim lsPage As String

    Dim lsTestID As String
    Dim lsSubCode As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    
    Dim lsresult_IFCC As String
    Dim lsresult_eAg As String
    
    
    Dim sSampleType As String
    Dim sLotNo As String
    Dim sLevel As String
    
    Dim rv As Integer
    Dim vTemp As String
    Dim qOrdDate As String
    Dim qQMCode As String
    Dim qOrdSeqNo As String
    Dim qEquipCode As String
    Dim qSpcCode As String
    Dim qExamCode As String
    Dim qSetYN As String
    Dim qLotNo As String
    Dim qRoomCode As String
    Dim qQCType As String
    Dim qEditID As String
    Dim qEditIP As String
    Dim qTransStr As String

    If asData = "" Then
        Exit Sub
    End If
    x = 0
    TablePtr = 1
    
'    For j = 1 To Len(asData)
'        If (Mid(asData, j, 1) = chrETX) Then
'            TablePtr = TablePtr + 1
'            ResultTbl(TablePtr) = " "
'        Else
'            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
'        End If
'    Next j
    
    Select Case Mid(asData, 2, 1)
    Case "H":       'Header Record
            Var_Clear
            gsSampleType = ""
            iCnt = 0
            
            For i = 1 To Len(asData)
                If Mid(asData, i, 1) = "|" Then
                    iCnt = iCnt + 1
    
                    Select Case iCnt
                        Case 11
                            gsSampleType = Mid(asData, i + 1, 1)
                        Case 13
                            gDate = Mid(asData, i + 1, 14)      '장비에서 받은 날짜시간
                    End Select
                End If
            Next i
    Case "P":
    Case "O":
            gsBarCode = Trim$(mGetP(ResultTbl(1), 4, "|"))
            gsPosNo = ""
            gsRackNo = ""
            gsSeqNo = ""
            
            gRow = -1
            For i = 1 To vasID.DataRowCnt
                If gsBarCode <> "" Then
                    If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                        gRow = i
                        Exit For
                    End If
    '            ElseIf sSampleType = "Q" Then
    
                End If
            Next i
            
            If gRow < 0 Then
                gRow = vasID.DataRowCnt + 1
                If vasID.maxrows < gRow Then
                    vasID.maxrows = gRow
                End If
            End If
            
            SetText vasID, gsBarCode, gRow, colBarcode
            SetText vasID, gsRackNo, gRow, colRack
            SetText vasID, gsPosNo, gRow, colPos
            
            vasActiveCell vasID, gRow, colBarcode
            ClearSpread vasRes
            
            '샘플정보 가져오기
            If gsSampleType = "Q" Then
                SetText vasID, "QC", gRow, colPName
            Else
                If Trim(GetText(vasID, gRow, colPID)) = "" And gsBarCode <> "" And Mid(gsBarCode, 1, 1) <> "U" Then
                    Get_Sample_Info gRow
                End If
            End If
    Case "R":
            gOrderMessage = "R"
            
    
            lsTestID = Trim$(mGetP(ResultTbl(1), 3, "|"))    '장비코드
            lsTestID = Trim$(mGetP(lsTestID, 4, "^"))    '장비코드
            lsResult = Trim$(mGetP(ResultTbl(1), 4, "|"))            '결과
            
            If lsTestID = "" Then: Exit Sub
            
            ClearSpread vasTemp
    
            SQL = "Select examcode, examname, seqno From equipexam" & vbCrLf & _
                  "Where equipno = '" & gEquip & "' " & vbCrLf & _
                  "And equipcode = '" & lsTestID & "' " ' & vbCrLf & _
                  "and examcode in (" & gOrderExam & ") "
            res = db_select_Col(gLocal, SQL)
            
            If res > 0 Then
                lsExamCode = Trim(gReadBuf(0))
                lsExamName = Trim(gReadBuf(1))
                lsSeqNo = Trim(gReadBuf(2))
                
                '숫자만 디스플레이 하기
                If IsNumeric(lsResult) = False Then
                    For ii = 1 To Len(lsResult)
                        If Mid(lsResult, ii, 1) = "?" Then
                            lsResult = Mid(lsResult, ii + 1)
                            
                            Exit For
                        End If
                    Next ii
                End If
                
                lsResRow = vasRes.DataRowCnt + 1
                If vasRes.maxrows < lsResRow Then
                    vasRes.maxrows = lsResRow
                End If
                
                '소수점 처리, 결과 형태 처리
                
                lsEquipRes = lsResult
                lsResult = SetResult(lsResult, lsTestID)
                lsResult_Buff = lsResult
                
                SetText vasRes, lsTestID, lsResRow, colEquipCode         '장비코드
                SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
                SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                SetText vasRes, lsResult, lsResRow, colResult            '결과
                
                SetText vasID, lsResult, gRow, colA1c                    '결과
                SetText vasID, gsFlag, gRow, colA1c + 1                  'Flag
                
                SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                SetText vasRes, gsFlag, lsResRow, 7                      'Flag
                
                
                Save_Local_One gRow, lsResRow, "1", CLng(lsEquipRes)
                            
                If IsNumeric(lsResult) = False Then
                    Exit Sub
                End If
    
                lsResult_Buff = ""
                    
            End If
    Case "L":
            gOrderExam = ""
            If MnTransAuto.Checked = True Then
                res = Insert_Data(gRow)
                
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
    End Select
    

    
End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    Dim strMsg As String
    Dim lngConnect  As Long
    
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
    
'    If Not Connect_Server Then
'        MsgBox "연결되지 않았습니다."
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If

'    '-- osw 추가
'    For i = 1 To 3
'        If Not Connect_PRServer Then
'            cn_cnt = cn_cnt + 1
'            If cn_cnt = 3 Then
'                If Not Connect_DRServer Then
'                    MsgBox "연결되지 않았습니다."
'                    cn_Server_Flag = False
'                    Exit Sub
'                Else
'                    cn_Server_Flag = True
'                End If
'            End If
'        Else
'            cn_Server_Flag = True
'        End If
'    Next

    lngConnect = dce_setenv(App.Path & "\sl.env", "", "")
    If lngConnect = 0 Then Call dce_error(strMsg)
    'MsgBox "성공"
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
    
    dtpFrDt.Value = Now
    dtpToDt.Value = Now + 1
    
    '==============================
    
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
    Dim intCnt As Integer
    
    With vasID
        For intCnt = 1 To .DataRowCnt
            .Col = 1
            .Row = intCnt
            If .Value = "1" Then
                Select Case intSndPhase
                    Case 1  '## Header
                        strOutput = intFrameNo & "H|\^&|||HOST|||||||P" & vbCr & ETX
                        intSndPhase = 2
                        intFrameNo = intFrameNo + 1
                    Case 2  '## Patient
                        strOutput = intFrameNo & "P|1||" & Trim(GetText(vasID, intCnt, 6)) & "||" & Trim(GetText(vasID, intCnt, 2)) & "l|" & vbCr & ETX
                        intSndPhase = 3
                        intFrameNo = intFrameNo + 1
                    Case 3  '## order
                        strOutput = intFrameNo & "O|1|" & Trim(GetText(vasID, intCnt, 3)) & "|" & Trim(GetText(vasID, intCnt, 4)) & "-" & Trim(GetText(vasID, intCnt, 5)) & _
                                                 "|^^^HB2CAP96|R||||||A||||||||||||||O" & vbCr & ETX
                        intSndPhase = 4
                        intFrameNo = intFrameNo + 1
                    Case 4  '## Termianator
                        strOutput = intFrameNo & "L|1|N" & vbCr & ETX
                        intSndPhase = 5
                        intFrameNo = intFrameNo + 1
                    Case 5  '## EOT
                        MSComm1.Output = EOT
                        Save_Raw_Data "[Tx]" & EOT
                        
                        Call Sleep(500)
                        .Col = 1
                        .Row = intCnt
                        .Value = "0"
                        
                        SetBackColor vasID, intCnt, intCnt, 1, colState, 234, 255, 154
                        SetText vasID, "Send", intCnt, colState

                        
                        intFrameNo = 1
                        intSndPhase = 1
                        
                        '-- 오더가 남아있으면
                        If intCnt < .DataRowCnt Then
                            MSComm1.Output = ENQ
                            Save_Raw_Data "[Tx]" & ENQ
                        Else
                            strState = ""
                        End If
                        
                        Exit Sub
                End Select
                
                strOutput = STX & strOutput & GetChkSum(strOutput) & vbCr & vbLf
                MSComm1.Output = strOutput
                Debug.Print strOutput
                Save_Raw_Data "[Tx]" & strOutput
                
                Exit For

            End If
        Next
    End With
    
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
            
''Rst:
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
                                If strState = "Q" Then Call SendOrder
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
        If vasID.maxrows < intRow Then
            vasID.maxrows = intRow
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
    
    intRow = -1
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colBarcode)) = pBarNo Then
            intRow = i
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasID.DataRowCnt + 1
        If vasID.maxrows < intRow Then
            vasID.maxrows = intRow
        End If
    End If
    
    Call SetText(vasID, pBarNo, intRow, colBarcode)  '3
    Call SetText(vasID, mResult.RackNo, intRow, colRack)       '4
    Call SetText(vasID, mResult.TubePos, intRow, colPos)         '5
    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    
    gRow = intRow
    

    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)
'
'    If Trim(strItems) = "" Then
'        mOrder.NoOrder = True
'        mOrder.Order = strItems
'    Else
'        mOrder.NoOrder = False
'        mOrder.Order = ""
'    End If
    

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
    
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        strType = Mid$(strRcvBuf, 2, 1)
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
'                strBarno = Format$(mGetP(strRcvBuf, 3, "|"), String$(10, "#"))
'                strTemp1 = mGetP(strRcvBuf, 4, "|")
'                strSeq = mGetP(strTemp1, 1, "^")
'                strRackNo = Format$(mGetP(strTemp1, 2, "^"), "####")
'                strTubePos = Format$(mGetP(strTemp1, 3, "^"), "##")
'
'                mResult.BarNo = strBarno
''                mResult.SpcPos = strTubePos & "/" & strRackNo
'                mResult.RackNo = strRackNo
'                mResult.TubePos = strTubePos
            
            Case "Q"    '## Request Information
                '## 바코드번호, SEQ, Disk No, Tube Position 조회
                If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
                strTemp1 = mGetP(strRcvBuf, 3, "|")
                strBarno = Trim$(mGetP(strTemp1, 2, "^"))
                
                mOrder.NoOrder = False
                mOrder.BarNo = strBarno
                mOrder.Seq = mGetP(strTemp1, 3, "^")
                mOrder.RackNo = mGetP(strTemp1, 4, "^")
                mOrder.TubePos = mGetP(strTemp1, 5, "^")
                
                Call GetOrder(strBarno)
                strState = "Q"
                
            Case "O"    '## Order
                strBarno = mGetP(strRcvBuf, 3, "|")
                strTemp1 = mGetP(strRcvBuf, 4, "|")
                'strSeq = mGetP(strTemp1, 1, "^")
                strRackNo = mGetP(strTemp1, 1, "-")
                strTubePos = mGetP(strTemp1, 2, "-")
                
                mResult.BarNo = strBarno
'                mResult.SpcPos = strTubePos & "/" & strRackNo
                mResult.RackNo = strRackNo
                mResult.TubePos = strTubePos
                
                Call SetPatInfo(strBarno)
                               

            Case "R"    '## Result
                '## 장비기준 검사명, 결과, Abnormal Flag
                strTemp1 = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                strTemp2 = mGetP(strRcvBuf, 4, "|")
                strFlag = mGetP(strRcvBuf, 7, "|")
                
                'strResult = strTemp2
                
                '## 정성, 정량결과를 동시에 수신할수 있도록 수정
                '## 정성, 정량에 따른결과처리, 결과에 "^"가 포함되면 정성결과
                
                '-- 결과 경우의 값
                '-- 1. Target Not Detected
                '-- 2. < 2.00E+1 IU/mL  ==> (2.00 * 5.82) * 10 & " 이하"  ==> 116 이하
                '-- 3. > 1.70E+8 IU/mL  ==> (1.70 * 5.82) & "X10^8 이상"  ==> 116 이하
                '-- 4. 6.24E+2 IU/mL
                
                
'                strTemp2 = Replace(strTemp2, "IU/mL", "")
'
'                If InStr(UCase(strTemp2), "TARGET") > 0 Then
'                    strResult = strTemp2
'                ElseIf UCase(strTemp2) = "FAILED" Then
'                    strResult = strTemp2
'                ElseIf InStr(strTemp2, "<") > 0 Then
'                    strTemp2 = Trim(Replace(strTemp2, "<", ""))
'                    strResult = Round(CDec(strTemp2) * calValue, 0)
'                    strResult = strResult & " 이하"
'                ElseIf InStr(strTemp2, ">") > 0 Then
'                    strTemp2 = Trim(Replace(strTemp2, ">", ""))
'                    strResult = mGetP(Trim(strTemp2), 1, "E+")
'                    strResult = Format(strResult * calValue, "##.000")
'                    strResult = strResult & " X 10^" & mGetP(Trim(strTemp2), 2, "+") & " 이상"
'                Else
'                    strResult = mGetP(Trim(strTemp2), 1, "E+")
'                    strResult = Format(strResult * calValue, "##.00")
'                    strResult = strResult & " X 10^" & mGetP(Trim(strTemp2), 2, "+")
'                End If
                strResult = strTemp2
                Select Case strFlag
                Case "N"
                    If InStr(UCase(strTemp2), "TARGET") > 0 Then
                        strResult = strTemp2
                    Else
                        strResult = CDec(strTemp2) * calValue

                        'strResult = strTemp2 / (10 ^ (InStr(strTemp2, ".") - 2))
                        'strResult = strResult * calValue
                        If strResult > 1000 Then
                            'strResult = Round(Mid(strResult, 1, 1) & "." & Mid(strResult, 2, 3), 2)
                            'strResult = Round(Mid(strResult, 1, 1) & "." & Mid(strResult, 2, 3), 2) & " X 10^" & Len(strResult) - 1
                            If InStr(strResult, ".") = 0 Then
                                strResult = Format(strResult, "##.0")
                            End If
                            strResult = Round(Mid(strResult, 1, 1) & "." & Mid(strResult, 2, 3), 2) & " X 10^" & Len(Mid(strResult, 1, InStr(strResult, ".") - 1)) - 1
                        Else
                            strResult = Round(strResult, 2)
                        End If
                        'strResult = Round(strResult, 2) & " X 10^" & InStr(strTemp2, ".") - 2
                        
                    End If
                Case "A"
                    If UCase(strResult) = "INVALID" Then  'Invalid
                        strResult = "9.894 X 10^8 이상"
                    Else
                        strResult = strTemp2 / (10 ^ (InStr(strTemp2, ".") - 2))
                        strResult = strResult * calValue
                        strResult = Round(strResult, 2) & " X 10^" & InStr(strTemp2, ".") - 2
                    End If
                Case ">"
'                    strResult = strTemp2 / (10 ^ 8)
'                    strResult = strResult * calValue
'                    strResult = Round(strResult, 2) & " X 10^" & InStr(strTemp2, ".") - 2 & " 이상"
                    strResult = "9.894 X 10^8 이상"
                Case "<"
                    'strResult = strResult / (10 ^ (InStr(strTemp2, ".") - 2))
                    'strResult = strResult * calValue
                    strTemp2 = Replace(strTemp2, "<", "")
                    strResult = CDec(strTemp2) * calValue
                    
                    strResult = Round(strResult, 0) & " 이하"
                
                End Select
                
                If strResult <> "" Then
                    '## 정성결과 저장
                    strIntBase = strTemp1
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
                        
                        lsResRow = vasRes.DataRowCnt + 1
                        If vasRes.maxrows < lsResRow Then
                            vasRes.maxrows = lsResRow
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
                        '-- 오더 없을 경우
                              SQL = "Select examcode, examname, seqno "
                        SQL = SQL & "  From equipexam"
                        SQL = SQL & " Where equipno = '" & gEquip & "' "
                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                        res = db_select_Col(gLocal, SQL)
                        
                        If res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.maxrows < lsResRow Then
                                vasRes.maxrows = lsResRow
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
                            strState = ""
                        End If
                    End If
                End If
                
                strState = "R"
                
            Case "C"    '## Comment
                 '## Abnormal 결과일때 Comment 저장
                If strFlag <> "N" Then
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    'strComm = "[Flag]: " & mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
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

                        'If IsNumeric(strResult) = False Then
                        '    Exit Sub
                        'End If
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


Sub VARIANTII(asData As String)
    
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim i As Integer
    Dim ii As Integer
    Dim j As Integer
    Dim k As Integer
    Dim x As Integer
    
    Dim iCnt As Integer
    
    Dim lsID As String
    Dim lsPid As String
    Dim lsPName As String
    Dim lsJumin1 As String
    Dim lsJumin2 As String
    Dim lsPSex As String
    Dim lsPage As String

    Dim lsTestID As String
    Dim lsSubCode As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    
    Dim lsresult_IFCC As String
    Dim lsresult_eAg As String
    
    
    Dim sSampleType As String
    Dim sLotNo As String
    Dim sLevel As String
    
    Dim rv As Integer
    Dim vTemp As String
    Dim qOrdDate As String
    Dim qQMCode As String
    Dim qOrdSeqNo As String
    Dim qEquipCode As String
    Dim qSpcCode As String
    Dim qExamCode As String
    Dim qSetYN As String
    Dim qLotNo As String
    Dim qRoomCode As String
    Dim qQCType As String
    Dim qEditID As String
    Dim qEditIP As String
    Dim qTransStr As String

    If asData = "" Then
        Exit Sub
    End If
    x = 0
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
    
    If Mid(ResultTbl(1), 2, 1) = "H" Then     'Header Record
        Var_Clear
        gsSampleType = ""
        iCnt = 0
        
        For i = 1 To Len(asData)
            If Mid(asData, i, 1) = "|" Then
                iCnt = iCnt + 1

                Select Case iCnt
                    Case 11
                        gsSampleType = Mid(asData, i + 1, 1)
                    Case 13
                        gDate = Mid(asData, i + 1, 14)      '장비에서 받은 날짜시간
                End Select
            End If
        Next i
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "O" Then
        If gsSampleType <> "P" Then: Exit Sub '/////QC데이터 안나와도 됨
        
        
        
        sTmp = Trim(ResultTbl(3))      'Barcode, Rack, Pos
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            If gsSampleType = "P" Then
                    If InStr(1, sTmp, "^") > 0 Then
                        iCnt = InStr(1, sTmp, "^")
                            gsBarCode = Trim(Mid(sTmp, 1, iCnt - 1))    'Barcode
                            If IsNumeric(gsBarCode) = True And Len(gsBarCode) > 12 Then
                                gsBarCode = Trim(Mid(gsBarCode, 1, 12))
                            End If
                        sTmp = Mid(sTmp, i + 1)
                        iCnt = InStr(1, sTmp, "^")
                            gsPosNo = Mid(sTmp, 1, iCnt - 1)       'Rack
                        sTmp = Mid(sTmp, 1)
                        iCnt = InStr(1, sTmp, "^")
                            gsRackNo = Mid(sTmp, iCnt + 1)     'pos
                    End If
'                If InStr(1, gsBarCode, "U") > 0 Then '////// Unknown 이 있을시에는
'                    gsBarCode = ""
'                End If
          
            ElseIf gsSampleType = "HC" Or gsSampleType = "LC" Then
                sLotNo = Trim(ResultTbl(16)) 'lotno
                i = InStr(1, sLotNo, "")
                If i > 0 Then
                    sLotNo = Mid(sLotNo, 1, i - 1)
                End If
                i = InStr(1, sLotNo, "^")
                If i > 0 Then
'                    sLevel = Mid(sLotNo, 1, i - 1)
'                    sLotNo = Mid(sLotNo, i + 1)
                    sLotNo = Mid(sLotNo, 1, i - 1)
                End If
            End If
        End If
        
        sTmp = Trim(ResultTbl(5))
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            i = InStr(1, sTmp, "^")
            sTmp = Mid(sTmp, i + 1)
            i = InStr(1, sTmp, "^")
            sTmp = Mid(sTmp, i + 1)
            i = InStr(1, sTmp, "^")
            gsSeqNo = Mid(sTmp, i + 1)
        End If
        
        
        
        
        gRow = -1
        For i = 1 To vasID.DataRowCnt
            If gsBarCode <> "" Then
                If Trim(GetText(vasID, i, colBarcode)) = gsBarCode Then
                    gRow = i
                    Exit For
                End If
'            ElseIf sSampleType = "Q" Then

            End If
        Next i
        
        If gRow < 0 Then
            gRow = vasID.DataRowCnt + 1
            If vasID.maxrows < gRow Then
                vasID.maxrows = gRow
            End If
        End If
        
        SetText vasID, gsBarCode, gRow, colBarcode
        SetText vasID, gsRackNo, gRow, colRack
        SetText vasID, gsPosNo, gRow, colPos
        
        vasActiveCell vasID, gRow, colBarcode
        ClearSpread vasRes
        
        '샘플정보 가져오기
        If gsSampleType = "Q" Then
            SetText vasID, "QC", gRow, colPName
        Else
            If Trim(GetText(vasID, gRow, colPID)) = "" And gsBarCode <> "" And Mid(gsBarCode, 1, 1) <> "U" Then
                Get_Sample_Info gRow
            End If
        End If
    End If
    
    
    If (Mid(ResultTbl(1), 2, 1) = "P") Then          'Test Order Record
        
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "L" Then
        If Trim(GetText(vasID, gRow, colPName)) <> "" Then
        
            gOrderExam = ""
            If MnTransAuto.Checked = True Then
                res = Insert_Data(gRow)
                
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
            
        End If
    SetText vasID, "Result", gRow, colState
    End If
    

    If (Mid(ResultTbl(1), 2, 1) = "R") Then     'Result
        gOrderMessage = "R"
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        lsTestID = Left(sTmp, i - 1)    '장비코드
        i = InStr(1, sTmp, "^")
        lsSubCode = Mid(sTmp, i + 1)
        sTmp = ResultTbl(4)
        lsResult = Trim(sTmp)           '결과
        
        
'        gsResDateTime = ResultTbl(10)    'result time
    
'        If Trim(gOrderExam) = "" Then
'            Exit Sub
'        End If
        If lsSubCode <> "AREA" Then: Exit Sub
        
        ClearSpread vasTemp

        SQL = "Select examcode, examname, seqno From equipexam" & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "And equipcode = '" & lsTestID & "' " ' & vbCrLf & _
              "and examcode in (" & gOrderExam & ") "
        res = db_select_Col(gLocal, SQL)
        
        If res > 0 Then
            lsExamCode = Trim(gReadBuf(0))
            lsExamName = Trim(gReadBuf(1))
            lsSeqNo = Trim(gReadBuf(2))
            
            '숫자만 디스플레이 하기
            If IsNumeric(lsResult) = False Then
                For ii = 1 To Len(lsResult)
                    If Mid(lsResult, ii, 1) = "?" Then
                        lsResult = Mid(lsResult, ii + 1)
                        
                        Exit For
                    End If
                Next ii
            End If
            
            lsResRow = vasRes.DataRowCnt + 1
            If vasRes.maxrows < lsResRow Then
                vasRes.maxrows = lsResRow
            End If
            
            '소수점 처리, 결과 형태 처리
            
            lsEquipRes = lsResult
            lsResult = SetResult(lsResult, lsTestID)
            lsResult_Buff = lsResult
            
            SetText vasRes, lsTestID, lsResRow, colEquipCode         '장비코드
            SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
            SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
            SetText vasRes, lsResult, lsResRow, colResult            '결과
            
            SetText vasID, lsResult, gRow, colA1c                    '결과
            SetText vasID, gsFlag, gRow, colA1c + 1                  'Flag
            
            SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
            SetText vasRes, gsFlag, lsResRow, 7                      'Flag
            
            
            Save_Local_One gRow, lsResRow, "1", CLng(lsEquipRes)
                        
            If IsNumeric(lsResult) = False Then
                Exit Sub
            End If
'//// IFCC, eAg 체크시
'''            For i = 1 To 2
'''                lsResRow = vasRes.DataRowCnt + 1
'''                If vasRes.MaxRows < lsResRow Then
'''                    vasRes.MaxRows = lsResRow
'''                End If
'''
'''                'IFCC,eAg 결과  처리
'''                If i = 1 Then
'''                    If gADD_IFCC = "-" Then
'''                        lsResult = CStr((CCur(gIFCC1) * CCur(lsResult_Buff)) - CCur(gIFCC2))
'''                    ElseIf gADD_IFCC = "+" Then
'''                        lsResult = CStr((CCur(gIFCC1) * CCur(lsResult_Buff)) + CCur(gIFCC2))
'''                    End If
'''                    lsResult = Format(lsResult, "####")
'''                    lsTestID = "IFCC"
'''                    lsExamCode = "B312002"
'''                    lsExamName = "IFCC"
'''                    lsSeqNo = "2"
'''                    lsResult = SetResult(lsResult, lsTestID)
'''                    SetText vasRes, lsResult, lsResRow, colResult           '결과
'''                    SetText vasID, lsResult, gRow, colIFCC              '결과
'''                    SetText vasID, gsFlag, gRow, colIFCC + 1          'Flag
'''                    SetText vasRes, gsFlag, lsResRow, 7          'Flag
'''                Else
'''                    If gADD_eAg = "-" Then
'''                        lsResult = CStr((CCur(geAg1) * CCur(lsResult_Buff)) - CCur(geAg2))
'''                    ElseIf gADD_eAg = "+" Then
'''                        lsResult = CStr((CCur(geAg1) * CCur(lsResult_Buff)) + CCur(geAg2))
'''                    End If
'''                    lsResult = Format(lsResult, "####")
'''                    lsTestID = "eAg"
'''                    lsExamCode = "B312003"
'''                    lsExamName = "eAg"
'''                    lsSeqNo = "3"
'''                    lsResult = SetResult(lsResult, lsTestID)
'''                    SetText vasRes, lsResult, lsResRow, colResult           '결과
'''                    SetText vasID, lsResult, gRow, coleAg               '결과
'''                    SetText vasID, gsFlag, gRow, coleAg + 1           'Flag
'''                    SetText vasRes, gsFlag, lsResRow, 7          'Flag
'''                End If
'''
'''                SetText vasRes, lsTestID, lsResRow, colEquipCode         '장비코드
'''                SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
'''                SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
'''                SetText vasRes, lsResult, lsResRow, colResult            '결과
'''                SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
'''
'''
'''                Save_Local_One gRow, lsResRow, "1"
'''            Next i
            
            lsResult_Buff = ""
                        
        End If
            
            
    End If
    
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
    
'    If IsNumeric(sEquipRes) = False Then
'        Exit Function
'    End If
    
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

Private Sub SSPanel1_DblClick()
    If FrmHideControl.Visible = True Then
        FrmHideControl.Visible = False
    Else
        FrmHideControl.Visible = True
    End If
    
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
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colSpecNo))
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' " & vbCrLf & _
          " AND RECENO = '" & Trim(GetText(vasID, Row, colRack)) & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "
    
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.maxrows = vasRes.DataRowCnt
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
    
    lsID = Trim(GetText(vasRID, Row, colSpecNo))
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
          " AND RECENO = '" & Trim(GetText(vasRID, Row, colRack)) & "' " & vbCrLf & _
          " AND EXAMDATE = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, REFFLAG , EQUIPRESULT"
    
    res = db_select_Vas(gLocal, SQL, vasRRes)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasRRes.maxrows = vasRRes.DataRowCnt
    
    For i = 1 To vasRRes.maxrows
        If Trim(GetText(vasRRes, i, colFLAG)) = "H" Then
            SetForeColor vasRRes, i, i, colResult, colResult, 255, 0, 0
        ElseIf Trim(GetText(vasRRes, i, colFLAG)) = "L" Then
            SetForeColor vasRRes, i, i, colResult, colResult, 0, 255, 0
        End If
    Next i
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
        vasRRes.maxrows = 0
        
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

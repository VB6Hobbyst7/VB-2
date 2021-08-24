VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   0  '없음
   Caption         =   " ROBO 888 Interface "
   ClientHeight    =   10680
   ClientLeft      =   330
   ClientTop       =   825
   ClientWidth     =   18825
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
   ScaleWidth      =   18825
   Begin VB.Frame Frame1 
      Height          =   9525
      Left            =   60
      TabIndex        =   40
      Top             =   750
      Width           =   15045
      Begin VB.CommandButton cmdPrint 
         Caption         =   "수동발행"
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
         Left            =   6630
         TabIndex        =   62
         Top             =   300
         Width           =   885
      End
      Begin VB.CommandButton cmdLocalList 
         Caption         =   "채혈내역"
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
         Left            =   5730
         TabIndex        =   57
         Top             =   300
         Width           =   885
      End
      Begin VB.CheckBox chkRef 
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   150
         TabIndex        =   55
         Top             =   330
         Value           =   1  '확인
         Width           =   225
      End
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   1815
         Left            =   8505
         TabIndex        =   48
         Top             =   6720
         Visible         =   0   'False
         Width           =   5970
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '평면
            Height          =   1455
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   49
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   660
         TabIndex        =   46
         Top             =   780
         Width           =   225
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
         Left            =   12000
         TabIndex        =   45
         Top             =   270
         Visible         =   0   'False
         Width           =   1395
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
         Left            =   13470
         TabIndex        =   44
         Top             =   270
         Width           =   1395
      End
      Begin VB.TextBox txtTest 
         Height          =   675
         Left            =   1680
         TabIndex        =   43
         Top             =   4800
         Visible         =   0   'False
         Width           =   4125
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   435
         Left            =   6060
         TabIndex        =   42
         Top             =   4950
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "조회"
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
         Left            =   4830
         TabIndex        =   41
         Top             =   300
         Width           =   885
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   8655
         Left            =   135
         TabIndex        =   47
         Top             =   720
         Width           =   14745
         _Version        =   393216
         _ExtentX        =   26009
         _ExtentY        =   15266
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
         MaxCols         =   20
         MaxRows         =   20
         MoveActiveOnFocus=   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":058D
         UserResize      =   2
      End
      Begin MSComCtl2.DTPicker dtpFrDt 
         Height          =   315
         Left            =   1740
         TabIndex        =   50
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   133365761
         CurrentDate     =   40739
      End
      Begin MSComCtl2.DTPicker dtpToDt 
         Height          =   315
         Left            =   3360
         TabIndex        =   51
         Top             =   330
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   133365761
         CurrentDate     =   40739
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "장비와 연결이 완료되었습니다."
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
         Height          =   225
         Left            =   8730
         TabIndex        =   61
         Top             =   480
         Width           =   3255
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   9420
         Picture         =   "frmInterface.frx":12DF
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   10560
         Picture         =   "frmInterface.frx":1869
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   11910
         Picture         =   "frmInterface.frx":1DF3
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port : "
         Height          =   180
         Index           =   0
         Left            =   8730
         TabIndex        =   60
         Top             =   210
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send : "
         Height          =   180
         Left            =   9885
         TabIndex        =   59
         Top             =   210
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Receive : "
         Height          =   180
         Left            =   10920
         TabIndex        =   58
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Ref"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   420
         TabIndex        =   56
         Top             =   390
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3180
         TabIndex        =   54
         Top             =   390
         Width           =   195
      End
      Begin VB.Label Label3 
         Caption         =   "조회기간"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   53
         Top             =   390
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "연결정보"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7620
         TabIndex        =   52
         Top             =   390
         Width           =   975
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   4665
      Left            =   14220
      TabIndex        =   19
      Top             =   5280
      Visible         =   0   'False
      Width           =   8175
      Begin VB.Timer Timer1 
         Left            =   2940
         Top             =   2070
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2070
         Top             =   2130
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   7200
         TabIndex        =   39
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
         SpreadDesigner  =   "frmInterface.frx":237D
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   180
         TabIndex        =   37
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
         TabIndex        =   26
         Top             =   1380
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   120
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   1320
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   960
         Left            =   6705
         TabIndex        =   21
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
         TabIndex        =   20
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
         SpreadDesigner  =   "frmInterface.frx":25A3
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   3195
         TabIndex        =   27
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
         SpreadDesigner  =   "frmInterface.frx":27C9
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   4980
         TabIndex        =   28
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
         SpreadDesigner  =   "frmInterface.frx":29EF
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   1800
         TabIndex        =   29
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
         SpreadDesigner  =   "frmInterface.frx":2C15
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   4860
         TabIndex        =   31
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   5700
         TabIndex        =   30
         Top             =   1410
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   2085
      Left            =   14340
      TabIndex        =   16
      Top             =   1950
      Visible         =   0   'False
      Width           =   9465
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   1260
         TabIndex        =   17
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
         SpreadDesigner  =   "frmInterface.frx":2E3B
      End
      Begin FPSpread.vaSpread vasPrintBuf 
         Height          =   1245
         Left            =   120
         TabIndex        =   18
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
         SpreadDesigner  =   "frmInterface.frx":48C2
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   9315
      Left            =   14910
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
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
      TabPicture(0)   =   "frmInterface.frx":4AE8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":4B04
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   8775
         Left            =   -74820
         TabIndex        =   7
         Top             =   360
         Width           =   14625
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   8460
            TabIndex        =   32
            Top             =   630
            Width           =   6015
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   38
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Left            =   4200
               TabIndex        =   36
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
               TabIndex        =   35
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Left            =   1605
               TabIndex        =   34
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
               TabIndex        =   33
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   240
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   240
            TabIndex        =   13
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
            Format          =   133562368
            CurrentDate     =   40457
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   780
            TabIndex        =   10
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
            TabIndex        =   9
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
            TabIndex        =   8
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   7815
            Left            =   240
            TabIndex        =   11
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
            SpreadDesigner  =   "frmInterface.frx":4B20
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   7275
            Left            =   8460
            TabIndex        =   12
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
            SpreadDesigner  =   "frmInterface.frx":553E
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10305
      Width           =   18825
      _ExtentX        =   33205
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
            TextSave        =   "2021-03-11"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오전 11:27"
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
      Caption         =   " ROBO 888 INTERFACE "
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
      Begin VB.Timer Timer2 
         Left            =   10980
         Top             =   270
      End
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7920
         Top             =   180
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7440
         Top             =   180
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   13710
         Picture         =   "frmInterface.frx":92E7
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   240
         Width           =   285
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   12600
         TabIndex        =   2
         Top             =   510
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
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
         Format          =   133562368
         CurrentDate     =   40457
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   6330
         Top             =   30
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
               Picture         =   "frmInterface.frx":9871
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":9E0B
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":A3A5
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":A93F
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":B1D1
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":B32B
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":B485
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCaution 
         BackColor       =   &H00C0FFFF&
         Caption         =   " 조회기간이 1일이 넘으면 병동별로 출력되지 않습니다"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   3810
         TabIndex        =   63
         Top             =   180
         Visible         =   0   'False
         Width           =   8325
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
         Index           =   1
         Left            =   11700
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
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
         Left            =   14010
         TabIndex        =   4
         Top             =   270
         Width           =   645
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
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "전송"
      Visible         =   0   'False
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

Dim strRoboRcvData  As String
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

Private Sub chkRef_Click()
    If chkRef.Value = "0" Then
        Timer1.Enabled = False
    Else
        Timer1.Interval = 5000
        Timer1.Enabled = True
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
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    'SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
'    vasRes.MaxRows = 0
    
'    dtptoday = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0
    
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

Private Sub cmdLocalList_Click()
        
    Call LoadRoboLocalData

End Sub

Private Sub cmdPrint_Click()
Dim intRow As Integer
Dim blnSnd As Boolean
    
    blnSnd = False
    With vasID
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = 1
            If .Value = "1" Then
                Call RoboMakeSend(intRow, "N")
                blnSnd = True
                .Col = 1
                .Value = "0"
            End If
        Next
        If blnSnd = True Then
            Call RoboMakeSend(intRow, "Y")
        End If
    End With

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

End Sub


Private Sub GetWorkList(ByVal pFrDt As String, ByVal pToDt As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strExamCode As String
    Dim rs As ADODB.Recordset
    Dim sSpecNo As String
    Dim strWardCode As String
    
    vasID.MaxRows = 0
    intRow = 0
    
    '-- 검사대상자 가져오기
          SQL = "Select WORK_DY,WORK_SQNO,SPCM_NO,TRMS_YN,TRMS_DT,EM_YN,CTNR_CD,SPCM_CD_NM," & vbCrLf
    SQL = SQL & "       PID , PT_NM, Sex, Age, MED_DVSN, MED_DP, BLCLR_NM, EXMN_NM, REGI_ID, RGST_DT, AMEN_ID, UPDT_DT " & vbCrLf
    SQL = SQL & "  From SPSLMIROB " & vbCrLf
    SQL = SQL & " Where WORK_DY between '" & Format(pFrDt, "yyyymmdd") & "' and '" & Format(pToDt, "yyyymmdd") & "'" & vbCrLf
    SQL = SQL & "   and TRMS_YN = 'N'"
    SQL = SQL & "   and TRMS_DT IS NULL"
    SQL = SQL & " Order By WORK_DY,WORK_SQNO "
    
    Set rs = cn_Ser.Execute(SQL, , 1)
          
    If Not rs.EOF Then
        Do Until rs.EOF
            '-- 바코드 번호 가져오기
            SQL = "SELECT FN_LABCVTPRTBCNO('" & Trim(rs.Fields("SPCM_NO")) & "') FROM DUAL "
            res = db_select_Col(gServer, SQL)
            sSpecNo = Trim(gReadBuf(0))
            '-- 부서명 조회
            SQL = "SELECT FN_GETDEPTCD('" & Trim(rs.Fields("MED_DP")) & "',TO_CHAR(SYSDATE,'yyyymmdd'),'0') FROM DUAL"
            res = db_select_Col(gServer, SQL)
            strWardCode = Trim(gReadBuf(0))
            '-- 2011.10.14 수정 : 입원(병동)/외래(부서) 코드는 장비에서 4자리만 받을수 있다.
            strWardCode = Mid(strWardCode, 1, 4)
            strWardCode = Replace(strWardCode, "/", "")
            intRow = intRow + 1
            vasID.MaxRows = intRow
            
            '-- 입원/외래
            If Trim(rs.Fields("MED_DVSN")) = "I" Then
                SetText vasID, "2", intRow, 2    '2:입원
            Else
                SetText vasID, "1", intRow, 2    '1:외래
            End If
            '-- 응급여부
            If Trim(rs.Fields("EM_YN")) = "Y" Then
                SetText vasID, "2", intRow, 3    '2:emergency
            Else
                SetText vasID, "1", intRow, 3    '1:normal
            End If
            '-- tube code
            SetText vasID, Mid(Trim(rs.Fields("CTNR_CD")), 2), intRow, 4    '01:PLAIN 6ml,02:SST 8ml,03:S.C,04:EDTA,99:other
            '-- work no
            SetText vasID, Format(Trim(rs.Fields("WORK_SQNO")), "0000"), intRow, 5
            '-- barcode no
            SetText vasID, sSpecNo, intRow, 6
            '-- ward code
            SetText vasID, strWardCode & Space$(4 - Len(strWardCode)), intRow, 7
            '-- 날짜
            SetText vasID, Format(Trim(rs.Fields("WORK_DY")), "####/##/##"), intRow, 8
            '-- 환자명
            SetText vasID, Trim(rs.Fields("PT_NM")) & Space$(14 - LengthByte(Trim(rs.Fields("PT_NM")))), intRow, 9
            '-- 환자ID
            SetText vasID, Trim(rs.Fields("PID")), intRow, 10
            '-- 환자 sex
            SetText vasID, Trim(rs.Fields("SEX")), intRow, 11
            '-- 환자age
            SetText vasID, Trim(rs.Fields("AGE")) & Space$(3 - Len(Trim(rs.Fields("AGE")))), intRow, 12
            '-- 검체코드명
            If Len(Trim(rs.Fields("SPCM_CD_NM"))) > 15 Then
                SetText vasID, Mid(Trim(rs.Fields("SPCM_CD_NM")), 1, 15), intRow, 13
            Else
                SetText vasID, Trim(rs.Fields("SPCM_CD_NM")) & Space$(15 - Len(Trim(rs.Fields("SPCM_CD_NM")))), intRow, 13
            End If
            '-- 채혈자
            SetText vasID, Trim(rs.Fields("BLCLR_NM")) & Space$(14 - LenB(Trim(rs.Fields("BLCLR_NM")))), intRow, 14
            '-- 검사명
            SetText vasID, Trim(rs.Fields("EXMN_NM")), intRow, 15
            '-- 검사명1
            SetText vasID, Mid(Trim(rs.Fields("EXMN_NM")), 1, 30), intRow, 16
            '-- 검사명2
            If Len(Trim(rs.Fields("EXMN_NM"))) > 30 And Len(Trim(rs.Fields("EXMN_NM"))) <= 60 Then
                SetText vasID, Mid(Trim(rs.Fields("EXMN_NM")), 31) & Space$(60 - Len(Trim(rs.Fields("EXMN_NM")))), intRow, 17
            ElseIf Len(Trim(rs.Fields("EXMN_NM"))) > 60 Then
                SetText vasID, Mid(Trim(rs.Fields("EXMN_NM")), 31), intRow, 17
            Else
                SetText vasID, Space$(30), intRow, 17
            End If
            
            '-- 검체수
                  SQL = "Select Count(*) " & vbCrLf
            SQL = SQL & "  From SPSLMIROB " & vbCrLf
            SQL = SQL & " Where TRMS_YN = 'N'"
            SQL = SQL & "   and TRMS_DT IS NULL"
            SQL = SQL & "   and WORK_DY = '" & Trim(rs.Fields("WORK_DY")) & "'"
            SQL = SQL & "   and PID = '" & Trim(rs.Fields("PID")) & "'"
            SQL = SQL & "   and CTNR_CD = '" & Trim(rs.Fields("CTNR_CD")) & "'"
            res = db_select_Col(gServer, SQL)
            
            SetText vasID, Trim(gReadBuf(0)), intRow, 18
            
            Call RoboMakeSend(intRow, "N")
            
            rs.MoveNext
        
        Loop
    
        Call RoboMakeSend(intRow, "Y")
        vasID.RowHeight(-1) = 12
    
'    Else
        
'        Call LoadRoboLocalData
    
    End If
    

End Sub

Private Sub RoboLocalSave(ByVal intRow As Integer, ByVal strMsg As String)
    Dim strTestNm As String
    
          SQL = "Delete From ROBO "
    SQL = SQL & " Where SPC_NO  = '" & Trim(GetText(vasID, intRow, 6)) & "' "
    cn.Execute SQL

          SQL = "Insert Into ROBO (SPC_NO,PT_NO,PT_NM,SEX,AGE,DEPT,SPC_NM,TEST_NM,ROBO_NO,ROBO_DAT,PH_NAME,SYS_DT,SYS_TM) "
    SQL = SQL & " Values ( "
    SQL = SQL & "'" & GetText(vasID, intRow, 6) & "', "
    SQL = SQL & "'" & GetText(vasID, intRow, 10) & "', "
    SQL = SQL & "'" & GetText(vasID, intRow, 9) & "', "
    SQL = SQL & "'" & GetText(vasID, intRow, 11) & "', "
    SQL = SQL & "'" & GetText(vasID, intRow, 12) & "', "
    SQL = SQL & "'" & GetText(vasID, intRow, 7) & "', "
    SQL = SQL & "'" & GetText(vasID, intRow, 13) & "', "
                    strTestNm = GetText(vasID, intRow, 15)
                    strTestNm = Replace(strTestNm, "'", "!")
    SQL = SQL & "'" & strTestNm & "', "
    SQL = SQL & "'" & GetText(vasID, intRow, 5) & "', "
                    'strMsg = Replace(strMsg, "'", "")
                    strMsg = Replace(strMsg, "'", "!")
    SQL = SQL & "'" & strMsg & "', "
    SQL = SQL & "'" & GetText(vasID, intRow, 14) & "', "
    SQL = SQL & "'" & Format(Now, "yyyymmdd") & "', "
    SQL = SQL & "'" & Format(Now, "hhmmss") & "') "
    
    cn.Execute SQL
    
End Sub

Private Sub RoboServerSaver(ByVal intRow As Integer, ByVal strMsg As String)

          SQL = "Update SPSLMIROB"
    SQL = SQL & "   Set TRMS_YN = 'Y',"
    SQL = SQL & "       TRMS_DT =  sysdate "
    SQL = SQL & " Where WORK_DY  = '" & Format(Trim(GetText(vasID, intRow, 8)), "yyyymmdd") & "' "
    SQL = SQL & "   And WORK_SQNO  = '" & Val(Trim(GetText(vasID, intRow, 5))) & "' "
    
    cn_Ser.Execute SQL
    
End Sub

Private Sub RoboMakeSend(ByVal intRow As Integer, Optional ByVal strEndFlag As String = "")
Dim strMsg As String
Dim strTestNm As String

    With vasID
        If strEndFlag = "N" Then
                     strMsg = "0"                           ' 1. message header(1)  0:text,9:end
            strMsg = strMsg & GetText(vasID, intRow, 2)     ' 2. in/out code(1)     1:외래,2:입원
            strMsg = strMsg & GetText(vasID, intRow, 3)     ' 3. label flag(1)      1:normal,2:emergency,3:dangil
            strMsg = strMsg & GetText(vasID, intRow, 4)     ' 4. tube code(2)       01:PLAIN 6ml,02:SST 8ml,03:S.C,04:EDTA,99:other
            strMsg = strMsg & GetText(vasID, intRow, 5)     ' 5. work no(4)         WORK_SQNO
            strMsg = strMsg & GetText(vasID, intRow, 6)     ' 6. barcode(10)
            strMsg = strMsg & GetText(vasID, intRow, 7)     ' 7. ward code(4)
            strMsg = strMsg & GetText(vasID, intRow, 8)     ' 8. date(10)           yyyy/mm/dd
            strMsg = strMsg & GetText(vasID, intRow, 9)     ' 9. patient name(14)
            strMsg = strMsg & GetText(vasID, intRow, 10)    '10. patient id(8)
            strMsg = strMsg & GetText(vasID, intRow, 11)    '11. sex(1)             M:남,F:여
            strMsg = strMsg & GetText(vasID, intRow, 12)    '12. age(3)
            strMsg = strMsg & GetText(vasID, intRow, 13)    '13. tube name(15)
            strMsg = strMsg & GetText(vasID, intRow, 14)    '14. 채혈자(14)
            strTestNm = GetText(vasID, intRow, 15)
            If LengthByte(strTestNm) > 60 Then
                strMsg = strMsg & Mid(strTestNm, 1, 60)  '15 검사명
            Else
                strMsg = strMsg & strTestNm & Space$(60 - LengthByte(strTestNm))  '15 검사명
            End If
            strMsg = strMsg & Format(GetText(vasID, intRow, 18), "00")   '17. 검체토탈갯수(2)
            strMsg = strMsg & Space$(50) & vbCr             '18. empty(50)
        
            If Winsock1.State = sckConnected Then
                imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
                If tmrSend.Enabled = False Then
                    tmrSend.Enabled = True
                Else
                    tmrSend.Enabled = False
                    tmrSend.Enabled = True
                End If
                
                Call Winsock1.SendData(strMsg)
                Save_Raw_Data "[Snd]" & strMsg
                
                Call RoboLocalSave(intRow, strMsg)
                Call RoboServerSaver(intRow, strMsg)
            End If
        Else
            strMsg = "9" & Space$(199) & vbCr
                
            If Winsock1.State = sckConnected Then
                imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
                If tmrSend.Enabled = False Then
                    tmrSend.Enabled = True
                Else
                    tmrSend.Enabled = False
                    tmrSend.Enabled = True
                End If
                
                Save_Raw_Data "[Snd]" & strMsg
                Call Winsock1.SendData(strMsg)

            End If
        End If
    End With
End Sub




Private Sub lblclear_Click()
    lblChangeBar.Caption = ""
    lblBarcode.Caption = ""
    lblChangePID.Caption = ""
    lblPname.Caption = ""
End Sub

Private Sub cmdRun()
    
    Dim itemX As ListItem
    
On Error GoTo ErrRoutine
'    CallForm = "frmInterface - Private Sub cmdRun()"
    
'    If Not cn_Local_Flag Then comEQP.PortOpen = True
    If cn_Local_Flag Then
'        Call ShowMessage("연결 되었습니다.")
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
        lblStatus = "작업중.."
    Else
'        Call ShowMessage("연결 되지 않았습니다.")
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
        lblStatus = "작업 대기중.."
    End If
        
Exit Sub
ErrRoutine:
'    Call ErrMsgProc(CallForm)
End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    Dim objSplash   As clsIISSplash
    Dim strProNm    As String           '프로그램 이름
    Dim strVer      As String           '프로그램 버전
    Dim strTitle    As String           'Title
    
    Screen.MousePointer = 11
            
    '## Splash화면표시
    strProNm = App.ProductName
    strVer = App.Major & "." & App.Minor & "." & App.Revision
    Set objSplash = New clsIISSplash
    With objSplash
        .ProjectNm = "ROBO 888" & Chr(13) & "자동채혈" & Chr(13) & "시스템" 'strProNm
        .Version = strVer
        .Message = strProNm & " 프로그램을 실행중입니다..."
        .LoadSplash
        DoEvents
    End With
    
    Call objSplash.SetMsg("프로그램 로딩중입니다...")
    DoEvents
            
    If App.PrevInstance Then
        Screen.MousePointer = 0
        End
    End If
    
    DoEvents
    
    '두번 실행 하지 않음
'    If App.PrevInstance Then
'       MsgBox "라벨러 프로그램이 이미 실행중입니다.", vbExclamation
'       End
'    End If
    
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    
    Me.Left = 0
    Me.Top = 0
    
    Me.Height = 11520
    Me.Width = 15435
    
    cmdIFClear_Click
    cmdRClear_Click
    lblclear_Click
    
    GetSetup
        
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Screen.MousePointer = 0
        Exit Sub
    Else
        cn_Local_Flag = True
    End If


    Call cmdRun                 ' 실행



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
    dtpToDt.Value = Now
    
    Winsock1.Protocol = sckTCPProtocol
    Winsock1.RemoteHost = gDRDB_Parm.ServerIP
    Winsock1.RemotePort = gDRDB_Parm.ServerPort
    Winsock1.Connect
    
    If Winsock1.State = sckConnected Then
        lblStatus.Caption = Winsock1.RemoteHost & ":" & Winsock1.RemotePort & " 에 연결되었습니다"
    ElseIf Winsock1.State = sckConnecting Then
        'Call LoadRoboLocalData
        lblStatus.Caption = Winsock1.RemoteHost & ":" & Winsock1.RemotePort & " 에 연결중.."
    End If
    
    Timer1.Interval = 5000
    Timer1.Enabled = True
    
'    Timer2.Interval = 1000
'    Timer2.Enabled = True
    '==============================
    
'    '-- osw 추가
'    For i = 1 To 3
'        If Not Connect_PRServer Then
'            cn_cnt = cn_cnt + 1
'            If cn_cnt = 3 Then
'                If Not Connect_DRServer Then
'                    MsgBox "연결되지 않았습니다."
'                    cn_Server_Flag = False
'                    Screen.MousePointer = 0
'                    Exit Sub
'                Else
                    cn_Server_Flag = True
'                End If
'            End If
'        Else
'            cn_Server_Flag = True
'        End If
'    Next
    
    Screen.MousePointer = 0

    
End Sub

Private Sub LoadRoboLocalData()
    Dim rs As ADODB.Recordset
    Dim intRow As Integer
    Dim strTmp As String
    Dim intCnt As Integer
    Dim intMinCnt As Integer
    Dim strTextNm As String
    
          SQL = "Select SPC_NO,PT_NO,PT_NM,SEX,AGE,DEPT,SPC_NM,TEST_NM,ROBO_NO,ROBO_DAT,PH_NAME,SYS_DT,SYS_TM "
    SQL = SQL & "  From ROBO    "
    SQL = SQL & " Where SYS_DT between '" & Format(dtpFrDt.Value, "yyyymmdd") & "' AND '" & Format(dtpToDt.Value, "yyyymmdd") & "'"
    SQL = SQL & "   and SPC_NO <> '' "
    SQL = SQL & " Order By SYS_DT,ROBO_NO,PT_NO,SPC_NM "
    
    Set rs = New ADODB.Recordset
    rs.Open SQL, cn
    
    With vasID
        .MaxRows = 0
        intRow = 0
        intMinCnt = 0
        If Not rs.EOF Then
            Do Until rs.EOF
                intRow = intRow + 1
                .MaxRows = intRow
                
                strTmp = Trim(rs.Fields("ROBO_DAT").Value)
                strTmp = Replace(strTmp, "!", "'")
                .SetText 2, intRow, Mid(strTmp, 2, 1)
                .SetText 3, intRow, Mid(strTmp, 3, 1)
                .SetText 4, intRow, Mid(strTmp, 4, 2)
                .SetText 5, intRow, Trim(rs.Fields("ROBO_NO").Value)
                .SetText 6, intRow, Trim(rs.Fields("SPC_NO").Value)
                .SetText 7, intRow, Mid(strTmp, 20, 4)
                .SetText 8, intRow, Mid(strTmp, 24, 10)
                
                'SetText vasID, Trim(rs.Fields("PT_NM")) & Space$(14 - LengthByte(Trim(rs.Fields("PT_NM")))), intRow, 9
                .SetText 9, intRow, Trim(rs.Fields("PT_NM").Value) & Space$(14 - LengthByte(Trim(rs.Fields("PT_NM"))))
                
                .SetText 10, intRow, Trim(rs.Fields("PT_NO").Value)
                .SetText 11, intRow, Trim(rs.Fields("SEX").Value)
                
'                SetText vasID, Trim(rs.Fields("AGE")) & Space$(3 - Len(Trim(rs.Fields("AGE")))), intRow, 12
                .SetText 12, intRow, Trim(rs.Fields("AGE").Value) & Space$(3 - Len(Trim(rs.Fields("AGE"))))
                
                            '-- 검체코드명
'                If Len(Trim(rs.Fields("SPCM_CD_NM"))) > 15 Then
'                    SetText vasID, Mid(Trim(rs.Fields("SPCM_CD_NM")), 1, 15), intRow, 13
'                Else
'                    SetText vasID, Trim(rs.Fields("SPCM_CD_NM")) & Space$(15 - Len(Trim(rs.Fields("SPCM_CD_NM")))), intRow, 13
'                End If
                If Len(Trim(rs.Fields("SPC_NM"))) > 15 Then
                    .SetText 13, intRow, Mid(Trim(rs.Fields("SPC_NM").Value), 1, 15)
                Else
                    .SetText 13, intRow, Trim(rs.Fields("SPC_NM")) & Space$(15 - Len(Trim(rs.Fields("SPC_NM"))))
                End If
                
'                '-- 채혈자
'                SetText vasID, Trim(rs.Fields("BLCLR_NM")) & Space$(14 - LenB(Trim(rs.Fields("BLCLR_NM")))), intRow, 14
                .SetText 14, intRow, Trim(rs.Fields("PH_NAME")) & Space$(14 - LenB(Trim(rs.Fields("PH_NAME"))))
                
                
'                '-- 검사명
'                SetText vasID, Trim(rs.Fields("EXMN_NM")), intRow, 15
                strTextNm = Trim(rs.Fields("TEST_NM").Value)
                strTextNm = Replace(strTextNm, "!", "'")
                .SetText 15, intRow, strTextNm
                
'                '-- 검사명1
'                SetText vasID, Mid(Trim(rs.Fields("EXMN_NM")), 1, 30), intRow, 16
                .SetText 16, intRow, Mid(strTextNm, 1, 30)
                
'                '-- 검사명2
'                If Len(Trim(rs.Fields("EXMN_NM"))) > 30 And Len(Trim(rs.Fields("EXMN_NM"))) <= 60 Then
'                    SetText vasID, Mid(Trim(rs.Fields("EXMN_NM")), 31) & Space$(60 - Len(Trim(rs.Fields("EXMN_NM")))), intRow, 17
'                ElseIf Len(Trim(rs.Fields("EXMN_NM"))) > 60 Then
'                    SetText vasID, Mid(Trim(rs.Fields("EXMN_NM")), 31), intRow, 17
'                Else
'                    SetText vasID, Space$(30), intRow, 17
'                End If
                If Len(strTextNm) > 30 And Len(strTextNm) <= 60 Then
                    .SetText 17, intRow, Mid(strTextNm, 31) & Space$(60 - Len(strTextNm))
                ElseIf Len(strTextNm) > 60 Then
                    .SetText 17, intRow, Mid(strTextNm, 31)
                Else
                    .SetText 17, intRow, Space$(30)
                End If
                
                .SetText 18, intRow, Mid(Right(strTmp, 53), 1, 2)
                .SetText 19, intRow, Trim(rs.Fields("SYS_DT").Value) & " " & Trim(rs.Fields("SYS_TM").Value)
                .SetText 20, intRow, "발행"
                
                rs.MoveNext
            Loop
        End If
        
        .RowHeight(-1) = 12
    End With
    
    Set rs = Nothing
    
End Sub

'문자열의 byte를 되돌려 준다.
Function LengthByte(ByVal Var As String) As Long
    Dim Cnt As Long
    Dim num As Long
    Dim TMP As String
    
    Cnt = 0: num = 0
    If Var = "" Then Exit Function
    Do
        Cnt = Cnt + 1: TMP = Mid(Var, Cnt, 1): num = num + 1
        If Asc(TMP) < 0 Then num = num + 1
    Loop Until Cnt >= Len(Var)
    LengthByte = num
End Function

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
            'strOutput = intFrameNo & "P|1|||||||||||||||||||||||||||||||||" & vbCr & ETX
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
                                
                                '3O|1|9905300211|1^00014^1^^SAMPLE^NORMAL|ALL|R|20110613090006|||||X||||||||||||||O|||||
                                '90
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


Private Sub Timer1_Timer()
    
    If Winsock1.State = sckConnected Then
        Call cmdWorkList_Click
    
    ElseIf Winsock1.State = sckClosed Then
        Winsock1.Protocol = sckTCPProtocol
        Winsock1.RemoteHost = gDRDB_Parm.ServerIP
        Winsock1.RemotePort = gDRDB_Parm.ServerPort
        Winsock1.Connect
'    ElseIf Winsock1.State = sckConnecting Then
'
'        Call LoadRoboLocalData
    End If
    
    If chkRef.Value = "0" Then
        Timer1.Enabled = False
    Else
        Timer1.Interval = 5000
        Timer1.Enabled = True
    End If

End Sub



Private Sub Timer2_Timer()
    If lblCaution.Visible = True Then
        lblCaution.Visible = False
    Else
        lblCaution.Visible = True
    End If
End Sub

Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
'    Dim i As Integer
'
'    For i = BlockRow To BlockRow2
'        vasID.Col = 1
'        vasID.Row = i
'        If vasID.Value = 0 Then
'        vasID.Value = 1
'        Else
'        vasID.Value = 0
'        End If
'    Next i
End Sub


Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
            
'        vasID_Click colBarcode, lRow
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

Private Sub Winsock1_Close()
    
    Winsock1.Close
    lblStatus.Caption = "장비와 연결이 종료 되었습니다."

End Sub

Private Sub Winsock1_Connect()
    
    lblStatus.Caption = "장비와 연결이 완료 되었습니다."

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strRcvData As String
    
    imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
    If tmrReceive.Enabled = False Then
        tmrReceive.Enabled = True
    Else
        tmrReceive.Enabled = False
        tmrReceive.Enabled = True
    End If
    
    Save_Raw_Data "[Rcv]" & strRoboRcvData
    Winsock1.GetData strRoboRcvData
    Debug.Print strRoboRcvData
    
    If strRoboRcvData = "00" Then
        '성공
    Else
        '실패
    End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    lblStatus.Caption = Number & ":" & Description
    Winsock1.Close

End Sub

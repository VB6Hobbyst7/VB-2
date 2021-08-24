VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmInterface 
   Caption         =   "Urometer720"
   ClientHeight    =   11100
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   20610
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
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   11100
   ScaleWidth      =   20610
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   9105
      Left            =   13770
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   10335
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   1815
         Left            =   8910
         TabIndex        =   79
         Top             =   3270
         Visible         =   0   'False
         Width           =   5970
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '평면
            Height          =   1455
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   80
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.CommandButton Command16 
         Caption         =   "전송테스트"
         Height          =   435
         Left            =   8310
         TabIndex        =   78
         Top             =   5880
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTest 
         Height          =   375
         Left            =   7620
         TabIndex        =   77
         Top             =   5910
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtNum 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3540
         TabIndex        =   76
         Text            =   "1"
         Top             =   6570
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton cmdPatDelete 
         Caption         =   "제외"
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
         Left            =   7440
         TabIndex        =   75
         Top             =   6750
         Width           =   795
      End
      Begin VB.ComboBox cboChk 
         Height          =   315
         ItemData        =   "frmInterface.frx":14F5
         Left            =   4230
         List            =   "frmInterface.frx":1502
         TabIndex        =   74
         Top             =   6480
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtBarNum 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5730
         TabIndex        =   73
         Top             =   6780
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   3375
         Left            =   180
         TabIndex        =   70
         Top             =   7320
         Visible         =   0   'False
         Width           =   8625
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1545
            Left            =   240
            TabIndex        =   71
            Top             =   1620
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
            SpreadDesigner  =   "frmInterface.frx":1518
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   1245
            Left            =   240
            TabIndex        =   72
            Top             =   300
            Width           =   8115
            _Version        =   393216
            _ExtentX        =   14314
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
            SpreadDesigner  =   "frmInterface.frx":2F91
         End
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "장비"
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
         Left            =   1515
         TabIndex        =   68
         Top             =   6450
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "수정"
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
         Left            =   2295
         TabIndex        =   67
         Top             =   6450
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   1455
         Left            =   120
         TabIndex        =   20
         Top             =   3270
         Width           =   3555
         _Version        =   393216
         _ExtentX        =   6271
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
         SpreadDesigner  =   "frmInterface.frx":31A9
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   2235
         Left            =   3780
         TabIndex        =   7
         Top             =   2550
         Width           =   4425
         _Version        =   393216
         _ExtentX        =   7805
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
         SpreadDesigner  =   "frmInterface.frx":33C1
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         Picture         =   "frmInterface.frx":35D9
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   21
         Top             =   5790
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   210
         TabIndex        =   19
         Top             =   5640
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
         Height          =   585
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   13
         Top             =   4830
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   5100
         TabIndex        =   12
         Top             =   5700
         Width           =   645
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
         Height          =   435
         Left            =   4440
         TabIndex        =   11
         Top             =   5715
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   3780
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   10
         Top             =   4830
         Visible         =   0   'False
         Width           =   4425
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
         Height          =   465
         Left            =   5820
         Style           =   1  '그래픽
         TabIndex        =   9
         Top             =   5640
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   4860
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   2835
         Begin VB.Timer tmrSend 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2220
            Top             =   300
         End
         Begin VB.Timer tmrReceive 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   1740
            Top             =   300
         End
         Begin MSCommLib.MSComm comEqp 
            Left            =   105
            Top             =   210
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
         Begin MSComctlLib.ImageList imlStatus 
            Left            =   1140
            Top             =   180
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
                  Picture         =   "frmInterface.frx":3B63
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":40FD
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":4697
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":4C31
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":54C3
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":561D
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":5777
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1485
         Left            =   120
         TabIndex        =   14
         Top             =   270
         Width           =   3585
         _Version        =   393216
         _ExtentX        =   6324
         _ExtentY        =   2619
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
         SpreadDesigner  =   "frmInterface.frx":58D1
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   2205
         Left            =   3780
         TabIndex        =   15
         Top             =   270
         Visible         =   0   'False
         Width           =   4395
         _Version        =   393216
         _ExtentX        =   7752
         _ExtentY        =   3889
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
         SpreadDesigner  =   "frmInterface.frx":5AE9
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1485
         Left            =   120
         TabIndex        =   16
         Top             =   1770
         Width           =   3585
         _Version        =   393216
         _ExtentX        =   6324
         _ExtentY        =   2619
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
         SpreadDesigner  =   "frmInterface.frx":5D01
      End
      Begin VB.Label lblDate 
         Caption         =   "2014-01-01"
         Height          =   165
         Left            =   630
         TabIndex        =   81
         Top             =   6900
         Width           =   1125
      End
      Begin VB.Label Label5 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "결과적용"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   450
         TabIndex        =   69
         Top             =   6540
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
         BorderStyle     =   1  '단일 고정
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
         Height          =   375
         Left            =   2010
         TabIndex        =   28
         Top             =   5760
         Width           =   825
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2940
         TabIndex        =   18
         Top             =   5730
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3720
         TabIndex        =   17
         Top             =   5730
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   20550
      TabIndex        =   29
      Top             =   0
      Width           =   20610
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Urometer720"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   210
         TabIndex        =   33
         Top             =   90
         Width           =   1200
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12180
         Picture         =   "frmInterface.frx":5F19
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13380
         Picture         =   "frmInterface.frx":64A3
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14670
         Picture         =   "frmInterface.frx":6A2D
         Top             =   90
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port"
         Height          =   195
         Index           =   0
         Left            =   11640
         TabIndex        =   32
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send"
         Height          =   195
         Left            =   12765
         TabIndex        =   31
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Receive"
         Height          =   195
         Left            =   13800
         TabIndex        =   30
         Top             =   120
         Width           =   735
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   10185
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   17965
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
      TabCaption(0)   =   "WorkList"
      TabPicture(0)   =   "frmInterface.frx":6FB7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "이전결과"
      TabPicture(1)   =   "frmInterface.frx":6FD3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   9675
         Left            =   -74820
         TabIndex        =   43
         Top             =   360
         Width           =   14625
         Begin VB.OptionButton optSaveResultR 
            Caption         =   "수정"
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
            Left            =   9735
            TabIndex        =   57
            Top             =   270
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton optSaveResultR 
            Caption         =   "장비"
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
            Left            =   8955
            TabIndex        =   56
            Top             =   270
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chkSave 
            Caption         =   "저장포함"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Left            =   6810
            TabIndex        =   55
            Top             =   210
            Value           =   1  '확인
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   7860
            TabIndex        =   49
            Top             =   630
            Width           =   6675
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   54
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   1
               Left            =   4590
               TabIndex        =   53
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
               Left            =   3540
               TabIndex        =   52
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   1
               Left            =   1995
               TabIndex        =   51
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label2 
               Caption         =   "챠트번호 :"
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
               Left            =   510
               TabIndex        =   50
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "&Excel"
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
            Left            =   13020
            TabIndex        =   48
            Top             =   -90
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
            Left            =   3720
            TabIndex        =   47
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   46
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
            Left            =   5250
            TabIndex        =   45
            Top             =   240
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdRClear 
            Caption         =   "&Clear"
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
            Left            =   13020
            TabIndex        =   44
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   8070
            Left            =   7860
            TabIndex        =   58
            Top             =   1455
            Width           =   6675
            _Version        =   393216
            _ExtentX        =   11774
            _ExtentY        =   14235
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
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":6FEF
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   59
            Top             =   300
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
            Format          =   21233664
            CurrentDate     =   40457
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   8805
            Left            =   150
            TabIndex        =   62
            Top             =   690
            Width           =   7605
            _Version        =   393216
            _ExtentX        =   13414
            _ExtentY        =   15531
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
            MaxCols         =   11
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":ACF1
            UserResize      =   2
         End
         Begin VB.Label Label7 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "결과적용"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   7890
            TabIndex        =   61
            Top             =   360
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label9 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   60
            Top             =   390
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         Height          =   9645
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   14625
         Begin VB.CheckBox chkSaveAll 
            Caption         =   "저장포함"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6060
            TabIndex        =   66
            Top             =   210
            Width           =   1125
         End
         Begin VB.CommandButton cmdDownload 
            Caption         =   "Down"
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
            Left            =   4890
            TabIndex        =   36
            Top             =   240
            Width           =   1035
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "조회"
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
            Left            =   3840
            TabIndex        =   37
            Top             =   240
            Width           =   1035
         End
         Begin VB.CommandButton cmdWorkDelete 
            Caption         =   "제외"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6480
            TabIndex        =   35
            Top             =   4680
            Width           =   975
         End
         Begin VB.CheckBox chkWAll 
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   420
            TabIndex        =   34
            Top             =   780
            Width           =   225
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "&Clear"
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
            Left            =   11940
            TabIndex        =   5
            Top             =   240
            Width           =   1275
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "&Save"
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
            Left            =   13260
            TabIndex        =   4
            Top             =   240
            Width           =   1245
         End
         Begin VB.Frame Frame6 
            Height          =   585
            Left            =   7740
            TabIndex        =   22
            Top             =   630
            Width           =   6825
            Begin VB.Label Label8 
               Caption         =   "등록번호 :"
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
               Left            =   270
               TabIndex        =   27
               Top             =   240
               Width           =   1230
            End
            Begin VB.Label lblBarcode 
               Caption         =   "12345"
               Height          =   165
               Index           =   0
               Left            =   1575
               TabIndex        =   26
               Top             =   240
               Width           =   585
            End
            Begin VB.Label Label6 
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
               Left            =   3360
               TabIndex        =   25
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   0
               Left            =   4410
               TabIndex        =   24
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   23
               Top             =   720
               Width           =   1155
            End
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   450
            TabIndex        =   3
            Top             =   5160
            Width           =   225
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   315
            Left            =   2460
            TabIndex        =   38
            Top             =   300
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   21233665
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   315
            Left            =   990
            TabIndex        =   39
            Top             =   300
            Width           =   1305
            _ExtentX        =   2302
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
            Format          =   21233665
            CurrentDate     =   40248
         End
         Begin FPSpread.vaSpread vasWorkList 
            Height          =   3855
            Left            =   90
            TabIndex        =   40
            Top             =   720
            Width           =   7605
            _Version        =   393216
            _ExtentX        =   13414
            _ExtentY        =   6800
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
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":B761
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8220
            Left            =   7740
            TabIndex        =   63
            Top             =   1260
            Width           =   6795
            _Version        =   393216
            _ExtentX        =   11986
            _ExtentY        =   14499
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
            MaxCols         =   8
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":C252
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   4425
            Left            =   90
            TabIndex        =   65
            Top             =   5070
            Width           =   7605
            _Version        =   393216
            _ExtentX        =   13414
            _ExtentY        =   7805
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
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":FF9B
            UserResize      =   2
         End
         Begin MSComCtl2.DTPicker dtpToday 
            Height          =   315
            Left            =   8730
            TabIndex        =   82
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
            Format          =   21233664
            CurrentDate     =   40457
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   7800
            TabIndex        =   83
            Top             =   390
            Width           =   780
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "검사진행"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   180
            TabIndex        =   64
            Top             =   4830
            Width           =   780
         End
         Begin VB.Label Label1 
            Appearance      =   0  '평면
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "처방일자"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label12 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2310
            TabIndex        =   41
            Top             =   360
            Width           =   105
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   10695
      Width           =   20610
      _ExtentX        =   36354
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   8123
            MinWidth        =   8114
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7938
            MinWidth        =   7938
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   6068
            MinWidth        =   6068
            TextSave        =   "2017-04-04"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "오전 9:58"
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
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "Setting"
      Begin VB.Menu MnTConfig 
         Caption         =   "통신설정"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "코드설정"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "Send"
      Begin VB.Menu MnTransAuto 
         Caption         =   "Auto"
         Checked         =   -1  'True
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "Manual"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Distinct '', SUBSTRING(O.OCMACPDTM,1,8)  AS 접수일시,'진료', R.RESOCMNUM  AS 내원번호, O.OCMCHTNUM  AS 차트번호,
'P.PBSPATNAM  AS 환자명, '','', P.PBSBIRDTE  AS 생년월일,'접수'"

'vasid, vasrid colum
Const colSpecNo = 0 '미사용
Const colCheckBox = 1
Const colOrdDate = 2    '검사의뢰일
Const colBarcode = 3    '챠트번호(=바코드)
Const colInNum = 4      '내원번호
Const colGumNum = 5     '검진번호
Const colSeqno = 6      'SeqNo
Const colRack = 7       'Rack
Const colPos = 8        'Pos
Const colPName = 9      '환자명
Const colPSex = 10      '성별
Const colOCnt = 11
Const colRCnt = 12
Const colState = 13


'seq 진료번호 검사의뢰일 바코드 진료과 구분 순번 접수번호 환자명 o r 상테

'Const colA1c = 12
'Const colIFCC = 13
'Const coleAg = 14




'sendflag
'0: Order
'1: Result
'2: Trans

'vasres, vasrres colum
Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
Const colMachResult = 4
Const colResult = 5
Const colSeq = 6
Const colFLAG = 7
Const colSubCode = 8

Dim gRow As Long

Dim gsBarCode       As String
Dim gsSampleType    As String
Dim gsPID           As String
Dim gsRackNo        As String
Dim gsPosNo         As String
Dim gsResDateTime   As String
Dim gsSeqNo         As String
Dim gsExamCode      As String
Dim gsExamName      As String
Dim gsOrder         As String
Dim gsResult        As String
Dim gsFlag          As String

Dim gMT             As String
Dim gComState       As Long
Dim gErrState       As Long

Dim strBuffer       As String
Dim strORQN         As String


'===============================
Const SPCLEN As Integer = 10

Const STX As String = ""
Const ETX As String = ""
Const ENQ As String = ""
Const ACK As String = ""
Const NAK As String = ""
Const EOT As String = ""
Const ETB As String = ""
Const FS  As String = ""
Const RS  As String = ""
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

Private Sub chkWAll_Click()
    Dim iRow As Long
    
    If chkWAll.Value = 1 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 1
        Next iRow
    ElseIf chkWAll.Value = 0 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 0
        Next iRow
    End If
End Sub
Private Sub cmdDownload_Click()
    Dim intRow As Integer
    Dim j  As Integer
    Dim varTmp As Variant
    
    j = 0
    With vasWorkList
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = colCheckBox
            If .Value = 1 Then
                vasID.MaxRows = vasID.MaxRows + 1
                
                SetText vasID, Trim(GetText(vasWorkList, intRow, colCheckBox)), vasID.MaxRows, colCheckBox
                SetText vasID, Trim(GetText(vasWorkList, intRow, colOrdDate)), vasID.MaxRows, colOrdDate
                SetText vasID, Trim(GetText(vasWorkList, intRow, colBarcode)), vasID.MaxRows, colBarcode
                SetText vasID, Trim(GetText(vasWorkList, intRow, colInNum)), vasID.MaxRows, colInNum
                SetText vasID, Trim(GetText(vasWorkList, intRow, colGumNum)), vasID.MaxRows, colGumNum
                SetText vasID, Trim(GetText(vasWorkList, intRow, colRack)), vasID.MaxRows, colRack
                SetText vasID, Trim(GetText(vasWorkList, intRow, colPos)), vasID.MaxRows, colPos
                SetText vasID, Trim(GetText(vasWorkList, intRow, colPName)), vasID.MaxRows, colPName
                SetText vasID, Trim(GetText(vasWorkList, intRow, colPSex)), vasID.MaxRows, colPSex

                txtNum = txtNum + 1
                
                .Col = colCheckBox
                .Value = "0"
                
            End If
        Next
    End With


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
            SetText vasPrint, Trim(GetText(vasRID, iRow, colBarcode)), j, 1
'            SetText vasPrint, Trim(GetText(vasRID, iRow, colPID)), j, 2
            SetText vasPrint, Trim(GetText(vasRID, iRow, colPName)), j, 3
'            SetText vasPrint, Trim(GetText(vasRID, iRow, colSex)), j, 4
            
            SQL = " SELECT RESULT " & vbCrLf & _
                  "   FROM PATRESULT " & vbCrLf & _
                  "  WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
                  "    AND EQUIPNO  = '" & gEquip & "' " & vbCrLf & _
                  "    AND BARCODE  = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf & _
                  "    AND PID      = '" & Trim(GetText(vasPrint, iRow, colBarcode)) & "' " & vbCrLf & _
                  "  ORDER BY SEQNO"
            Res = GetDBSelectVas(gLocal, SQL, vasPrintBuf)
            
            'sA1c = GetText(vasPrintBuf, 1, 1)
            'sIFCC = GetText(vasPrintBuf, 2, 1)
            'seAg = GetText(vasPrintBuf, 3, 1)

            ClearSpread vasPrintBuf, 1, 1

            'SetText vasPrint, sA1c, j, 7
            'SetText vasPrint, sIFCC, j, 8
            'SetText vasPrint, seAg, j, 9
            
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
    
    SetForeColor vasWorkList, 1, vasWorkList.MaxRows, 1, vasWorkList.MaxCols, 0, 0, 0
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    vasWorkList.MaxRows = 0
'    dtptoday = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0
    
End Sub

Private Sub cmdIFTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            
            Res = SaveTransDataW(lRow)
        
            If Res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "Failed", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "Trans", lRow, colState
                
                SQL = " UPDATE PATRESULT SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasID, lRow, colBarcode)) & "' "
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
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

Private Sub cmdPatDelete_Click()
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With vasWorkList
        For i = .DataRowCnt To 1 Step -1
            .Row = i
            .Col = colCheckBox
            If .Value = "1" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
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
    
    'SELECT 처음 '' 는 체크박스
          SQL = " SELECT '', BARCODE, DISKNO, POSNO, PID, PNAME,PSEX, PAGE, COUNT(*), COUNT(*), SENDFLAG " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    If chkSave.Value = "1" Then
        SQL = SQL & "    AND SENDFLAG IN ('0','1','2') " & vbCrLf
    Else
        SQL = SQL & "    AND SENDFLAG IN ('0','1') " & vbCrLf
    End If
    SQL = SQL & "  GROUP BY BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, SENDFLAG "
          
    Res = GetDBSelectVas(gLocal, SQL, vasRID)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For iRow = 1 To vasRID.DataRowCnt
        Select Case Trim(GetText(vasRID, iRow, colState))
            Case "0": SetText vasRID, "에러", iRow, colState
            Case "1": SetText vasRID, "결과", iRow, colState
            Case "2": SetText vasRID, "완료", iRow, colState
                      SetBackColor vasRID, iRow, iRow, 1, colState, 202, 255, 112
        End Select
    Next iRow
    
End Sub

Private Sub cmdRTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasRID.DataRowCnt
        vasRID.Row = lRow
        vasRID.Col = 1
        If vasRID.Value = 1 Then
            Res = SaveTransDataR(lRow)
        
            If Res = -1 Then
                SetForeColor vasRID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasRID, "Failed", lRow, colState
            ElseIf Res = 0 Then
            
            Else
                vasRID.Row = lRow
                vasRID.Col = 1
                vasRID.Value = 1
                
                SetBackColor vasRID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasRID, "Trans", lRow, colState
                
                SQL = " UPDATE PATRESULT SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquipCode & "' " & vbCrLf & _
                      " AND BARCODE = '" & Trim(GetText(vasRID, lRow, colBarcode)) & "' "
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
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



Private Sub GetWorkList(ByVal pFrDt As String, ByVal pToDt As String, Optional pBarNo As String)
    Dim i           As Integer
    Dim j           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strExamCode As String
    Dim RS          As ADODB.Recordset
    Dim sSpecNo     As String
    Dim buff        As String
    Dim strTestNm   As String
    Dim strDate     As String
    Dim strChart    As String
    Dim blnSame     As Boolean
    
    If pBarNo = "" Then
        vasWorkList.MaxRows = 0
        intRow = 0
    End If
    
    blnSame = False
    
    '-- BIT
          SQL = " SELECT DISTINCT '1', SUBSTRING(O.OCMACPDTM,1,8) AS 접수일시,O.OCMCHTNUM AS 차트번호,R.RESOCMNUM AS 내원번호,'', '','','', P.PBSPATNAM AS 환자명, P.PBSSEXTYP AS 성별" & vbCrLf
    SQL = SQL & "   FROM RESINF AS R, OCMINF AS O, PBSINF AS P, LABMST AS E, ODRINF AS W" & vbCrLf
    SQL = SQL & " WHERE O.OCMACPDTM BETWEEN '" & pFrDt & "000000" & "' AND '" & pToDt & "235959" & "'" & vbCrLf
    SQL = SQL & "   AND O.OCMCOMSTT NOT IN ('CN', 'CR', 'VC')" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD IN (" & gAllExam & ")" & vbCrLf
    SQL = SQL & "   AND R.RESOCMNUM = O.OCMNUM" & vbCrLf
    SQL = SQL & "   AND O.OCMCHTNUM = P.PBSCHTNUM" & vbCrLf
    SQL = SQL & "   AND R.RESOCMNUM = W.ODROCMNUM" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD *= W.ODRCOD" & vbCrLf
    SQL = SQL & "   AND R.RESLABCOD *= E.LABCOD" & vbCrLf
    '-- 저장미포함
    If chkSaveAll.Value = "0" Then
        SQL = SQL & "   AND (R.RESREPTYP IS NULL OR R.RESREPTYP <> 'F') " & vbCrLf         '--  'I':중간 'F' 완료"
        SQL = SQL & "   AND W.ODRDELFLG = 'N'" & vbCrLf
        SQL = SQL & "   AND (R.RESRLTVAL = ''  OR R.RESRLTVAL IS NULL)" & vbCrLf
    End If
    SQL = SQL & " ORDER BY 접수일시, 차트번호, 내원번호"
    
    SetRawData "[워크조회]" & SQL
    
    Res = GetDBSelectVas(gServer_Bit, SQL, vasWorkList)
    
    
'''          SQL = "SELECT Distinct '', A.REQUEST_DATE AS 접수일시, A.CHART_NO AS 차트번호,'',A.EXAM_NO AS 검진번호,'','','',A.PERSON_NAME AS 환자명,"
'''    SQL = SQL & "       CASE WHEN substring(A.PERSONAL_ID, 7, 1) = '1' THEN 'M' WHEN substring(A.PERSONAL_ID, 7, 1) = '3' THEN 'M' WHEN substring(A.PERSONAL_ID, 7, 1) = '2' THEN 'F' WHEN substring(A.PERSONAL_ID, 7, 1) = '4' THEN 'F' ELSE '' END AS '성별'"
'''    SQL = SQL & "  FROM MEDITOLISS..TOTAL A, MEDITOLISS..TOTRES B"
'''    SQL = SQL + " WHERE A.REQUEST_DATE between '" & pFrDt & "' And '" & pToDt & "'" & vbCr
'''    SQL = SQL & "   And B.EXAM_CODE in (" & gAllExam & ") " & vbCrLf
'''    SQL = SQL & "   AND B.EXAM_PART = 'C'"
'''    '-- 저장미포함
'''    If chkSaveAll.Value = "0" Then
'''        SQL = SQL & "   AND B.RESULT_VALUE = ''"
'''    End If
'''    SQL = SQL & "   AND A.REQUEST_DATE = B.REQUEST_DATE"
'''    SQL = SQL & "   AND A.EXAM_NO = B.EXAM_NO"
'''
''''    Res = GetDBSelectVas(gServer_Bit, SQL, vasID)
'''
'''    Set RS = cn_Ser.Execute(SQL, , 1)
'''
'''    If RS.EOF = True Or RS.BOF = True Then
'''        'MsgBox dtpStartDt.Value & "일 에서 " & dtpStopDt.Value & "일 까지" & vbNewLine & "조회된 결과가 없습니다." & vbNewLine & vbNewLine & "처방일자를 확인해주세요", vbOKOnly + vbInformation, Me.Caption
'''        'dtpStartDt.SetFocus
'''        Exit Sub
'''    End If
'''
'''    Do Until RS.EOF
'''        With vasWorkList
'''            For i = 1 To .DataRowCnt
'''                strDate = GetText(vasWorkList, i, colOrdDate)
'''                strChart = GetText(vasWorkList, i, colBarcode)
'''
'''                If Trim(RS("접수일시")) = strDate And Trim(RS("차트번호")) = strChart Then
'''                    SetText vasWorkList, Trim(RS.Fields("검진번호")) & "", i, colGumNum
'''                    blnSame = True
'''                End If
'''            Next
'''
'''            If blnSame = False Then
'''                .MaxRows = .MaxRows + 1
'''
'''                SetText vasWorkList, "1", .MaxRows, colCheckBox
'''                SetText vasWorkList, Trim(RS.Fields("접수일시")) & "", .MaxRows, colOrdDate
'''                SetText vasWorkList, Trim(RS.Fields("차트번호")) & "", .MaxRows, colBarcode
'''                SetText vasWorkList, Trim(RS.Fields("검진번호")) & "", .MaxRows, colGumNum
'''                SetText vasWorkList, Trim(RS.Fields("환자명")) & "", .MaxRows, colPName
'''                SetText vasWorkList, Trim(RS.Fields("성별")) & "", .MaxRows, colPSex
'''            End If
'''
'''            blnSame = False
''''            If pBarNo <> "" Then
''''                vasWorkList.MaxRows = vasWorkList.MaxRows + 1
''''                intRow = vasWorkList.MaxRows
''''            Else
''''                intRow = intRow + 1
''''                vasWorkList.MaxRows = intRow
''''            End If
'''
'''        End With
'''
'''        RS.MoveNext
'''    Loop
'''
'''    For i = 1 To vasWorkList.DataRowCnt
'''        Call vasWorkList.SetText(colRack, i, "0")
'''        Call vasWorkList.SetText(colPos, i, txtNum)
'''        txtNum = txtNum + 1
'''    Next
    
    vasWorkList.RowHeight(-1) = 12

    
End Sub

Private Sub cmdSearch_Click()
                
    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    
    vasID.RowHeight(-1) = 12

End Sub

Private Sub cmdWorkDelete_Click()
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With vasID
        For i = .DataRowCnt To 1 Step -1
            .Row = i
            .Col = colCheckBox
            If .Value = "1" Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                j = j + 1
            End If
        Next
    End With
End Sub


'Private Sub Form_Resize()
'    On Error Resume Next
'
'    If frmInterface.ScaleHeight = 0 Then Exit Sub
'
'
''    If cmdSL.Caption = "▶" Then
''        Frame1.Height = frmInterface.ScaleHeight - (Picture2.Top) - 1200
''        vasID.Height = Frame1.Height - 300
''
''        Frame1.Width = frmInterface.ScaleWidth - 200
''        vasID.Width = frmInterface.ScaleWidth - 7300
''
''
''        Frame6.Left = vasID.Width + 300
''        vasRes.Height = vasID.Height - 550
''        vasRes.Left = Frame6.Left
''    Else
'        stInterface.Width = Picture1.Width
'
'        Frame1.Height = frmInterface.ScaleHeight - (Picture1.Top) - 1200
'        vasID.Height = Frame1.Height - 300
'
'        Frame1.Width = frmInterface.ScaleWidth - 200
'        vasID.Width = frmInterface.ScaleWidth - 300
'
'        'Frame6.Left = frmInterface.ScaleWidth - vasID.Width
'        'vasRes.Height = vasID.Height - 550
'        'vasRes.Left = Frame6.Left
'
''    End If
'
'    Picture1.Width = Frame1.Width
'
'    StatusBar1.Panels(3).Width = Frame1.Width - 8500
'End Sub

Private Sub imgPort_DblClick()
    
'    '-- 개발시에만 Remark 풀어서 테스트진행
'    If FrmHideControl.Visible = True Then
'        Me.Width = 15435
'        FrmHideControl.Visible = False
'    Else
'        Me.Width = 25000
'        FrmHideControl.Visible = True
'    End If

End Sub

Private Sub lblclear_Click()
    lblChangePID.Caption = ""
    lblChangeBar.Caption = ""
    lblBarcode(0).Caption = ""
    lblBarcode(1).Caption = ""
    lblPname(0).Caption = ""
    lblPname(1).Caption = ""
End Sub

Private Sub Command16_Click()
    Dim i As Long
    Dim lsChar As String
        
    strBuffer = strBuffer & "00762" & vbCr
    strBuffer = strBuffer & "� RESULT  " & vbCr
    strBuffer = strBuffer & "p 75" & vbCr
    strBuffer = strBuffer & "q 07/09/07 10h16mn07s" & vbCr
    strBuffer = strBuffer & "u 0000000000000001" & vbCr
    strBuffer = strBuffer & "s 0002" & vbCr
    strBuffer = strBuffer & "v                               " & vbCr
    strBuffer = strBuffer & "t R" & vbCr
    strBuffer = strBuffer & "� D" & vbCr
    strBuffer = strBuffer & "! 007.9  " & vbCr
    strBuffer = strBuffer & "2 04.76  " & vbCr
    strBuffer = strBuffer & "3 014.0  " & vbCr
    strBuffer = strBuffer & "4 040.4  " & vbCr
    strBuffer = strBuffer & "5 084.9  " & vbCr
    strBuffer = strBuffer & "6 029.5  " & vbCr
    strBuffer = strBuffer & "7 034.7  " & vbCr
    strBuffer = strBuffer & "8 014.5  " & vbCr
    strBuffer = strBuffer & "@ 00263  " & vbCr
    strBuffer = strBuffer & "A 006.7  " & vbCr
    strBuffer = strBuffer & "B  .175  " & vbCr
    strBuffer = strBuffer & "C 014.1  " & vbCr
    strBuffer = strBuffer & "# 024.6  " & vbCr
    strBuffer = strBuffer & "% 009.7  " & vbCr
    strBuffer = strBuffer & "' 065.7  " & vbCr
    strBuffer = strBuffer & Chr(22) & " 001.9" & vbCr
    strBuffer = strBuffer & "$ 000.7  " & vbCr
    strBuffer = strBuffer & "& 005.3  " & vbCr
'    strBuffer = strBuffer & "W           *.4027Fa댉得�琿江뻻yiZQJFBB@DHKMJHBB?=94555222.****)))))))))*.000002799;@JOUUX\ejrwy{�냸냺룗몡룘뱫룕몡땲냸?tpig_\XZ\ "
'    strBuffer=strBuffer & "X             !! !!!!!!""%&+-7>Pbw뫃켈醇�埰威�땋�?yl_WTICA=732.,+(++('''*%'''))''&&&%$&&#$$#"""""!"!!!       !                  '
'    strBuffer=strBuffer & "Y           #);KVW^aqs�|냵?냲?uufciaaPLDGDC;;6610,-.210,,))&'''&&#$"$&&&$$""&&$! !!#"" !!! !""! !"!!!!""#!!!! """   !""""""$""!!
    strBuffer = strBuffer & "S       " & vbCr
    strBuffer = strBuffer & "_ 105" & vbCr
    strBuffer = strBuffer & "P           G3" & vbCr
    strBuffer = strBuffer & "] 000 000 000 027 039" & vbCr
    strBuffer = strBuffer & "?LC 250  " & vbCr
    strBuffer = strBuffer & "?V2.8 " & vbCr
    strBuffer = strBuffer & "?A75C" & vbCr
    strBuffer = strBuffer & "   " & vbCr

    strBuffer = STX & ";" & ETX & vbCr
    
    strBuffer = ":N1    80 81                 00620141422      15 1   7.0  2   4.1  3   0.5  4   4.5  5    34  6    20  7   417  8   239  9    97 14    85 15    14 16   0.7 18    93 19      T54     1 "
    
    Call comEqp_OnComm
        

End Sub



Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    'Me.Height = 11520
    Me.Width = 15435
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    cmdIFClear_Click
    cmdRClear_Click
    lblclear_Click
    
    GetSetup
    
    frmInterface.StatusBar1.Panels(1).Text = gUserID
        
    cboChk.ListIndex = 0
    
    comEqp.CommPort = gSetup.gPort
    comEqp.RTSEnable = gSetup.gRTSEnable
    comEqp.DTREnable = gSetup.gDTREnable
    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If
    
    If comEqp.PortOpen Then
        frmInterface.StatusBar1.Panels(2).Text = "연결 되었습니다"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        lblStatus = "작업중.."
    Else
        frmInterface.StatusBar1.Panels(2).Text = "연결 되지 않았습니다"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
'        lblStatus = "작업 대기중.."
    End If

    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    '-- osw 추가
    For i = 1 To 1
        If Not Connect_PRServer Then
            MsgBox "연결되지 않았습니다."
            cn_Server_Flag = False
            Exit Sub
        Else
            cn_Server_Flag = True
        End If
    Next
    
    '-- osw 추가
    For i = 1 To 1
        If Not Connect_DRServer Then
            MsgBox "연결되지 않았습니다."
            cn_Server_Flag = False
            Exit Sub
        Else
            cn_Server_Flag = True
        End If
    Next
    
    GetExamCode
    
    SetExamCode
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -30), "yyyymmdd")
    
    SQL = "delete from PATRESULT where examdate < '" & sDate & "'"
    Res = SendQuery(gLocal, SQL)
    
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
    
    '-- test
'    vasID.MaxRows = 10
    
End Sub



Function GetExamCode() As Integer
'    Dim i, j As Long
'
'    ClearSpread vasTemp
'    GetExamCode = -1
'    gAllExam = ""
'    SQL = "Select equipcode, examcode, subcode, examname, resprec, seqno " & vbCrLf & _
'          "  From EQPMASTER " & vbCrLf & _
'          " Where equipno = '" & gEquip & "' " & vbCrLf & _
'          " Order by  examcode "
'    Res = GetDBSelectVas(gLocal, SQL, vasCode)
'    If Res > 0 Then
'        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 7)
'    Else
'        SaveQuery SQL
'        Exit Function
'    End If
'
'    For i = 1 To vasCode.DataRowCnt
'        If i = 1 Then
'            gAllExam = "'" & Trim(GetText(vasCode, i, 2)) & Trim(GetText(vasCode, i, 3)) & "'"
'        Else
'            gAllExam = gAllExam & ",'" & Trim(GetText(vasCode, i, 2)) & Trim(GetText(vasCode, i, 3)) & "'"
'        End If
'
'        gArrEquip(i, 1) = i
'        For j = 1 To 5
'            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
'        Next j
'
'
'    Next i
'
'    GetExamCode = 1

    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    gAllExam = ""
    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
          "  From EQPMASTER " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by  examcode "
    Res = GetDBSelectVas(gLocal, SQL, vasCode)
    If Res > 0 Then
        ReDim gArrEquip(1 To vasCode.DataRowCnt, 1 To 7)
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
        For j = 1 To 6
            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
        Next j
        
        
    Next i
    
    GetExamCode = 1
End Function

Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 8
            
'            cboTest.AddItem gArrEquip(i + 1, 4) & Space(20) & "|" & gArrEquip(i + 1, 3)
        Next
        
'        cboTest.ListIndex = 0
    End With
    
    With vasWorkList
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            Call SetText(vasWorkList, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 8
        Next
    End With
        
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If comEqp.PortOpen = True Then
        comEqp.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Server
    DisConnect_Local
    Unload Me
    End
End Sub

Private Sub MnExamConfig_Click()
    'frmTestSet.Show
    frmTestSet.Show
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
    
    '-- ASTM TYPE별 Define 해야함.
    '-- ASTM TYPE = Standard
    If gASTMFormat = "1" Then
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
                    strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & "^" & mOrder.TubePos & "^^SAMPLE^NORMAL|ALL" & _
                                "|R||||||C||||||||||||||Q" & vbCr & ETX
                    intSndPhase = 5
                
                Else
                    If mOrder.IsSending = False Then   '## 최초 보낼때
                        strOutput = "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & "^" & mOrder.TubePos & "^^SAMPLE^NORMAL|" & mOrder.Order & "|R||||||N||||||||||||||Q"
                        
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
                comEqp.Output = EOT
                SetRawData "[Tx]" & EOT
                intFrameNo = 1
                
                Exit Sub
        End Select
    '-- ASTM TYPE = Long [=VISTA 500, Hitachi, Modular]
    ElseIf gCOMFormat = "2" Then
        Select Case intSndPhase
            Case 0
                strOutput = EOT
                comEqp.Output = strOutput
                'Save_Raw_Data "[Tx]" & strOutput
                strState = ""
                Exit Sub
    
            Case 1  '## Header
                '## Header
                strOutput = "H|\^&||||||||||P|" & vbCr
    
                '## Patient
                strOutput = strOutput & "P|1|" & vbCr
    
                '## Order
                If mOrder.NoOrder = False Then
                    strOutput = strOutput & "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||S||||||||||Q" & vbCr
                    
                    Select Case gOPTVersion
                    Case "1.0"  '## Version 1.0
                                'strOutput = strOutput & "O|1|0^" & Format$(mOrder.BarNo, String$(13, "@")) & "^" & mOrder.SpcType & "^" & mOrder.RackNo & "^" & mOrder.Pos & "|" & _
                                                        mOrder.Kind & "|" & mOrder.GetOrder & "|" & mOrder.Priority & "||||||N||^^||||||^^^^||||||O" & vbCr
                    Case "1.3"  '## Version 1.3
                                'strOutput = strOutput & "O|1|" & mOrder.BarNo & "|" & mOrder.GetInstSpcId & "|" & mOrder.GetOrder & "|" & mOrder.GetPriority & _
                                                        "||||||N||||" & mOrder.GetSampleType & "||||||||||O" & vbCr
                    End Select
                
                Else
                    '## 접수정보가 없는경우: 검사항목 정보를 보내지 않음!
                    strOutput = strOutput & "O|1|" & mOrder.BarNo & "|||R||||||C||||||||||||||Q" & vbCr
                    
                    Select Case gOPTVersion
                    Case "1.0"  '## Version 1.0
                                'strOutput = strOutput & "O|1|0^" & Format$(mOrder.BarNo, String$(13, "@")) & "^" & mOrder.SpcType & "^" & mOrder.RackNo & "^" & mOrder.Pos & "|" & _
                                                        mOrder.Kind & "||" & mOrder.Priority & "||||||N||^^||||||^^^^||||||O" & vbCr
                    Case "1.3"  '## Version 1.3
                                'strOutput = strOutput & "O|1|" & mOrder.BarNo & "|" & mOrder.GetInstSpcId & "||R" & _
                                                        "||||||N||||" & mOrder.GetSampleType & "||||||||||O" & vbCr
                    End Select
                
                End If
    
                '## Termianator
                strOutput = strOutput & "L|1|N" & vbCr
                strOutput = intFrameNo & strOutput
            Case 2
    
        End Select
    
        If Len(strOutput) >= 230 Then
            mOrder.Order = Mid$(strOutput, 231)
            strOutput = Mid$(strOutput, 1, 230) & ETB
            intSndPhase = 2
        Else
            strOutput = strOutput & ETX
            intSndPhase = 0
        End If
    
    End If
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    comEqp.Output = strOutput
    Debug.Print strOutput
    SetRawData "[Tx]" & strOutput
    
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

Private Sub tmrReceive_Timer()
    
    imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrReceive.Enabled = False

End Sub

Private Sub tmrSend_Timer()
    
    imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
    tmrSend.Enabled = False

End Sub

Public Sub SndMore()
    Dim strSndMsg As String
    
    'Call Sleep(1000)
    
    strSndMsg = ">"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) ' & GetChkSum(strSndMsg) & vbCr
    comEqp.Output = strSndMsg & vbCrLf
    
    'SetRawData "[Tx]" & strSndMsg & vbCrLf
    Debug.Print "[SndMore]" & strSndMsg
    
End Sub

Public Sub SndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & vbCrLf
    
End Sub


Private Sub comEqp_OnComm()
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

            Buffer = comEqp.Input
            'Buffer = strBuffer
            
            'txtData = txtData & Buffer
            
            SetRawData "[Rx]" & Buffer
            lngBufLen = Len(Buffer)
            
            Debug.Print Buffer

'''            For i = 1 To lngBufLen
'''                BufChar = Mid$(Buffer, i, 1)
'''                Select Case BufChar
'''                    Case STX
'''                        strBuffer = ""
'''
'''                    Case vbLf
'''                        Call EditRcvDataASTM
'''                        strBuffer = ""
'''
'''                    Case Else
'''                        strBuffer = strBuffer & BufChar
'''                End Select
'''            Next i

            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case intPhase
                    Case 1
                        Select Case BufChar
                            Case "~"
                                strBuffer = ""
                                intPhase = 2
                            Case Else
                                strBuffer = strBuffer & BufChar
                        End Select
                    Case 2
                    
                        Select Case BufChar
                            Case "~"
                                Call EditRcvDataASTM
                                strBuffer = ""
                                intPhase = 1
                            Case Else
                                strBuffer = strBuffer & BufChar
                        End Select
                End Select
            Next i


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

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, 표시, 검사오더만들기
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
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
    
    '-- 장비수신정보 표시
    Call SetText(vasID, mOrder.Seq, intRow, colSeqno)       ' Seq번호
    Call SetText(vasID, mOrder.RackNo, intRow, colRack)     ' Rack번호
    Call SetText(vasID, mOrder.TubePos, intRow, colPos)     ' Pos번호
    
    '-- 환자정보 표시
    strDate = Trim(GetText(vasWorkList, gRow, colOrdDate))
    strInNum = Trim(GetText(vasWorkList, gRow, colInNum))
    strGumNum = Trim(GetText(vasWorkList, gRow, colGumNum))
    
    Call SetText(vasID, strDate, intRow, colOrdDate)                                    ' 접수일자
    Call SetText(vasID, pBarNo, intRow, colBarcode)                                     ' 바코드(=챠트번호)
    Call SetText(vasID, strInNum, intRow, colInNum)                                     ' 내원번호
    Call SetText(vasID, strGumNum, intRow, colGumNum)                                   ' 검진번호
    Call SetText(vasID, Trim(GetText(vasWorkList, gRow, colPName)), intRow, colPName)   ' 이름
    Call SetText(vasID, Trim(GetText(vasWorkList, gRow, colPSex)), intRow, colPSex)     ' 성별
    
    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    
    '-- 검사자 정보 서버테이블에서 가져와 표시(for 워크리스트)  '6,7,8,9
    'Call GetSampleInfoW(intRow)
    
    '-- 바코드번호에 존재하는 검사코드 가져오기(인수 : 장비코드,챠트번호,접수일,내원번호,검진번호)
    gOrderExam = GetOrderExamCode(gEquip, pBarNo, strDate, strInNum, strGumNum)
    
    '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
    'strItems = GetGetEquipExamCode_Hitachi7080(gEquip, pBarNo, intRow)

    '-- 검사채널로 장비오더 만들기
'    If Trim(strItems) = "" Then
'        mOrder.NoOrder = True
'        mOrder.Order = strItems
'    Else
'        mOrder.NoOrder = False
'        mOrder.Order = strItems
'    End If
    
    Call SetText(vasID, "Order", intRow, colState)         '12 진행상태

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
    Dim strTestDt   As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
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
    
    '-- 장비수신정보 표시
    Call SetText(vasID, "1", intRow, colCheckBox)
    'Call SetText(vasID, pBarNo, intRow, colBarcode)
    Call SetText(vasID, mResult.RackNo, intRow, colRack)
    Call SetText(vasID, mResult.TubePos, intRow, colPos)
    Call vasActiveCell(vasID, intRow, colBarcode)
    
    '-- 접수일자,내원번호, 검진번호
    strDate = Trim(GetText(vasID, intRow, colOrdDate))
    strInNum = Trim(GetText(vasID, intRow, colInNum))
    strGumNum = Trim(GetText(vasID, intRow, colGumNum))
    
    '-- 결과스프레드 지우기
    Call ClearSpread(vasRes)
    
    '-- 검사자 정보 서버테이블 가져와 표시(for 워크리스트)  '5,6,7,8
    'Call GetSampleInfoW(intRow)
    
    '-- 현재 Row
    gRow = intRow
    
    '-- 검사자 정보 서버테이블에서 가져와 표시(for 워크리스트)  '6,7,8,9
    'Call GetSampleInfoW(intRow)
    
    '-- 바코드번호에 존재하는 검사코드 가져오기(인수 : 장비코드,챠트번호,접수일,내원번호,검진번호)
    gOrderExam = GetOrderExamCode(gEquip, pBarNo, strDate, strInNum, strGumNum)
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarno     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strFunc      As String
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    Dim strDate      As String
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsSubCode As String
    
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim i As Integer
    'Dim varORQN As Variant
    
    'Dim varH7080CH As Variant
    Dim strSendData As String
    Dim strSndMsg   As String
    Dim strExamCode() As String
    Dim fLCURO As Variant
    
    fLCURO = Split(strBuffer, vbCrLf)
    For intCnt = 1 To UBound(fLCURO)
        strRcvBuf = fLCURO(intCnt)
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        Select Case intCnt
            Case 3
                    strSeq = Mid(strRcvBuf, 10)
                    strSeq = Replace(strSeq, ")", "")
                    strSeq = Replace(strSeq, "(", "")
                    strSeq = Val(Trim(strSeq))
                    strBarno = ""
                    
                    If Len(strSeq) > 0 Then
                        For i = 1 To vasID.DataRowCnt
                            vasID.Row = i
                            vasID.Col = colCheckBox
                            If vasID.Value = "1" Then
                                gRow = i
                                strBarno = GetText(vasID, i, colBarcode)
                                Exit For
                            End If
                        Next i
                    End If
                    
                    If strBarno = "" Then
                        strBarno = strSeq
                    End If
                    
                    With mResult
                        .BarNo = strBarno
                    End With
                    
                    Call SetPatInfo(strBarno)
                    
                    strState = "O"
                
            Case 4 To 13
            
                    strIntBase = Mid(strRcvBuf, 1, 4)
                    strIntBase = Trim(strIntBase)
                    
                    strResult = Mid(strRcvBuf, 8, 4) '-- 정성
                    strResult = Trim(strResult)
                        
                    If Trim(strResult) = "-" Then
                        strResult = "Neg"
                    ElseIf Trim(strResult) = "+" Then
                        strResult = "1+"
                    ElseIf Trim(strResult) = "++" Then
                        strResult = "2+"
                    ElseIf Trim(strResult) = "+++" Then
                        strResult = "3+"
                    ElseIf Trim(strResult) = "++++" Then
                        strResult = "4+"
                    ElseIf UCase(Trim(strResult)) = "NORM" Then
                        strResult = "Neg"
                    ElseIf UCase(Trim(strResult)) = "+/-" Then
                        strResult = "Trace"
                    End If
                    
                    If strIntBase = "pH" Or strIntBase = "S.G" Then
                        strResult = Trim(Mid(strRcvBuf, 4))  '-- 정량
                        strResult = Replace(strResult, ">", "")
                        strResult = Replace(strResult, "<", "")
                        strResult = Replace(strResult, "=", "")
                    End If
                                            
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        SQL = SQL & "   AND EXAMCODE IN (" & gOrderExam & ") "
                        
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        '-- 오더 있을 경우
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '소수점 처리, 결과 형태 처리
                            lsEquipRes = strResult
                            'strResult = SetResult(strResult, strIntBase)
                            'lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText vasID, "Result", gRow, colState                 '11 진행상태
    
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                            SetText vasRes, strResult, lsResRow, colResult          '결과
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                            'SetText vasRes, lsSubCode, lsResRow, colSubCode         'SUBCODE
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- 로컬 저장
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            
                            strState = "R"
                            
                        '-- 오더 없을 경우
                        Else
                        
                                  SQL = "Select EXAMCODE,EXAMNAME,SEQNO "
                            SQL = SQL & "  From EQPMASTER"
                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                            Res = GetDBSelectColumn(gLocal, SQL)
                            
                            If Res > 0 Then
                                lsExamCode = Trim(gReadBuf(0))
                                lsExamName = Trim(gReadBuf(1))
                                lsSeqNo = Trim(gReadBuf(2))
                                
                                lsResRow = vasRes.DataRowCnt + 1
                                If vasRes.MaxRows < lsResRow Then
                                    vasRes.MaxRows = lsResRow
                                End If
                                
                                '소수점 처리, 결과 형태 처리
                                lsEquipRes = strResult
                                'strResult = SetResult(strResult, strIntBase)
                                'lsResult_Buff = strResult
                                
                                '-- Work List
                                SetText vasID, "Result", gRow, colState                 '진행상태
                                
                                '-- 결과 List
                                SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                                SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                                SetText vasRes, strResult, lsResRow, colResult          '결과
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                'SetText vasRes, lsSubCode, lsResRow, colSubCode         'SUBCODE
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                '-- 로컬 저장
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                
                                strState = "R"
                                
                            End If
                        End If
                    End If
            Case 14
                    '## DB에 결과저장
                    If MnTransAuto.Checked = True And strState = "R" Then
                        
                        SetText vasID, "", gRow, colCheckBox
                        
                        Res = SaveTransDataW(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
                            SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                            SetText vasID, "Failed", gRow, colState
                        Else
                            '-- 저장 성공
                            SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                            SetText vasID, "Trans", gRow, colState
                            
                                  SQL = "Update PATRESULT Set sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gEquip & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
                            SQL = SQL & "   And EXAMDATE = '" & strDate & "' "
                            
                            Res = SendQuery(gLocal, SQL)
                            
                            If Res = -1 Then
                                SaveQuery SQL
                                Exit Sub
                            End If
                        End If
                    End If
                    
                
                    strState = ""
        End Select
    Next

End Sub


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataGAS()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarno     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과(정성)
    Dim strIntResult As String   '수신한 결과(정량)
    Dim strQCResult  As String   '수신한 결과(QC)
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
    Dim strTmp      As String
    Dim intIdx      As Integer
    Dim intRow      As Integer
    
    Dim varBuffer   As Variant
    
    varBuffer = Split(strBuffer, vbCr)
    
    For intCnt = 1 To UBound(varBuffer)
        strRcvBuf = varBuffer(intCnt)
        
        Select Case intCnt
            Case 17    '## Sample No.  3            Syringe
                strSeq = Val(Mid(strRcvBuf, 11, 3))
                intRow = -1
                For ii = 1 To vasID.DataRowCnt
                    If Trim(GetText(vasID, ii, colSeqno)) = strSeq Then
                        intRow = ii
                        gRow = intRow
                        vasID.Row = gRow
                        vasID.Col = colBarcode
                        strBarno = Trim(vasID.Text)
                        Call SetPatInfo(strBarno)
                        Exit For
                    End If
                Next
            
                If strSeq = "" Then Exit Sub
                
                strState = "O"
            
            Case 18 To 49 '## Result
                '## 장비기준 검사명, 결과, Abnormal Flag
                strIntBase = intCnt
                strResult = Trim(Mid(strRcvBuf, 1, 4))
                
                If strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                    
                    Res = GetDBSelectColumn(gLocal, SQL)
                    
                    '-- 오더 있을 경우
                    If Res > 0 Then
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
                        
                        '-- Work List
                        SetText vasID, "Result", gRow, colState                 '11 진행상태
                        

                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                        SetText vasRes, strResult, lsResRow, colResult          '결과
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- 로컬 저장
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                        
                    '-- 오더 없을 경우
                    Else
                    
                              SQL = "Select examcode, examname, seqno "
                        SQL = SQL & "  From EQPMASTER"
                        SQL = SQL & " Where equipno = '" & gEquip & "' "
                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        If Res > 0 Then
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
                            
                            '-- Work List
                            SetText vasID, "Result", gRow, colState                 '진행상태
                            
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
                            SetText vasRes, strResult, lsResRow, colResult          '결과
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- 로컬 저장
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            
                        End If
                    End If
                End If
            Case 50   '## Terminator
                '## DB에 결과저장
                If MnTransAuto.Checked = True And strState = "R" Then
                        
                    Res = SaveTransDataW(gRow)
                    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                        SetText vasID, "Failed", gRow, colState
                    Else
                        '-- 저장 성공
                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                        SetText vasID, "Trans", gRow, colState
                        
                        SQL = " Update PATRESULT Set " & vbCrLf & _
                              " sendflag = '2' " & vbCrLf & _
                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
                              " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
                        Res = SendQuery(gLocal, SQL)
                        If Res = -1 Then
                            SaveQuery SQL
                            Exit Sub
                        End If
                    End If
                End If
            
                'SetText vasID, "Result", gRow, colState
                strState = ""
        
        End Select
    Next

End Sub
'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
'Private Sub EditRcvDataAU()
'    Dim strRcvBuf    As String   '수신한 Data
'    Dim strType      As String   '수신한 Record Type
'    Dim strBarno     As String   '수신한 바코드번호
'    Dim strSeq       As String   '수신한 Sequence
'    Dim strRackNo    As String   '수신한 Rack Or Disk No
'    Dim strTubePos   As String   '수신한 Tube Position
'    Dim strIntBase   As String   '수신한 장비기준 검사명
'    Dim strResult    As String   '수신한 결과
'    Dim strQCResult  As String   '수신한 결과(QC)
'    Dim strFlag      As String   '수신한 Abnormal Flag
'    Dim strComm      As String   '수신한 Comment
'    Dim strTemp1     As String
'    Dim strTemp2     As String
'    Dim intCnt       As Integer
'
'    Dim lsExamCode As String
'    Dim lsExamName As String
'    Dim lsSeqNo As String
'    Dim lsResult_Buff As String
'    Dim lsExamDate As String
'    Dim lsEquipRes As String
'    Dim lsResRow    As String
'    Dim ii As Integer
'    Dim strTmp      As String
'    Dim intIdx      As Integer
'
'
'    For intCnt = 1 To UBound(strRecvData)
'        strRcvBuf = strRecvData(intCnt)
'        strType = Mid$(strRcvBuf, 1, 2)
'
'        Select Case strType
'            '## Order Begin =========================================
'            Case "RB"   '## Begin Inquiry Text
'            Case "R "    '## Inquiry Order
'                strBarno = Trim(Mid(strRcvBuf, 14, 20))
'                strRackNo = Mid(strRcvBuf, 3, 4)
'                strTubePos = Mid(strRcvBuf, 7, 2)
'
'                With mOrder
'                    .BarNo = strBarno
'                    .RackNo = strRackNo
'                    .TubePos = strTubePos
'                    .Seq = Mid(strRcvBuf, 9, 5)
'                End With
'
'                Call GetOrder(strBarno)
'
'            Case "RE"   '## End Inquirty Text
'
'            '## Result =========================================
'            Case "DB"   '## Begin Result Text
'            Case "D "    '## Result
'                strBarno = Trim$(Mid$(strRcvBuf, 14, 10))
'
'                With mResult
'                    .BarNo = strBarno
'                    .RackNo = Mid(strRcvBuf, 3, 4)
'                    .TubePos = Mid(strRcvBuf, 7, 2)
'                End With
'
'                If strBarno = "" Then Exit Sub
'
'                strTmp = Mid$(strRcvBuf, 29)
'
'                Call SetPatInfo(strBarno)
'
'                Do While Len(strTmp) >= 11
'                    strIntBase = Mid$(strTmp, 2, 2)
'                    strResult = Mid$(strTmp, 4, 6)
'                    strComm = Mid$(strTmp, 10, 1)
'
'                    If strResult <> "" Then
'                        SQL = ""
'                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'                        SQL = SQL & "  FROM EQPMASTER"
'                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
'
'                        Res = GetDBSelectColumn(gLocal, SQL)
'
'                        '-- 오더 있을 경우
'                        If Res > 0 Then
'                            lsExamCode = Trim(gReadBuf(0))
'                            lsExamName = Trim(gReadBuf(1))
'                            lsSeqNo = Trim(gReadBuf(2))
'
'                            lsResRow = vasRes.DataRowCnt + 1
'                            If vasRes.MaxRows < lsResRow Then
'                                vasRes.MaxRows = lsResRow
'                            End If
'
'                            '소수점 처리, 결과 형태 처리
'                            lsEquipRes = strResult
'                            strResult = SetResult(strResult, strIntBase)
'                            lsResult_Buff = strResult
'
'                            '-- Work List
'                            SetText vasID, "Result", gRow, colState                 '11 진행상태
'
'
'                            '-- 결과 List
'                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
'                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
'                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
'                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'                            SetText vasRes, strResult, lsResRow, colResult          '결과
'                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
'                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'                            '-- 로컬 저장
'                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
'
'                            lsResult_Buff = ""
'
'                            strState = "R"
'
'                        '-- 오더 없을 경우
'                        Else
'
'                                  SQL = "Select examcode, examname, seqno "
'                            SQL = SQL & "  From EQPMASTER"
'                            SQL = SQL & " Where equipno = '" & gEquip & "' "
'                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
'                            Res = GetDBSelectColumn(gLocal, SQL)
'
'                            If Res > 0 Then
'                                lsExamCode = Trim(gReadBuf(0))
'                                lsExamName = Trim(gReadBuf(1))
'                                lsSeqNo = Trim(gReadBuf(2))
'
'                                lsResRow = vasRes.DataRowCnt + 1
'                                If vasRes.MaxRows < lsResRow Then
'                                    vasRes.MaxRows = lsResRow
'                                End If
'
'                                '소수점 처리, 결과 형태 처리
'                                lsEquipRes = strResult
'                                strResult = SetResult(strResult, strIntBase)
'                                lsResult_Buff = strResult
'
'                                '-- Work List
'                                SetText vasID, "Result", gRow, colState                 '진행상태
'
'                                '-- 결과 List
'                                SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
'                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
'                                SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
'                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '장비결과
'                                SetText vasRes, strResult, lsResRow, colResult          '결과
'                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
'                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
'                                '-- 로컬 저장
'                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
'
'                                lsResult_Buff = ""
'
'                            End If
'                        End If
'                    End If
'                    strTmp = Mid$(strTmp, 12)
'                Loop
'
'
'                If MnTransAuto.Checked = True And strState = "R" Then
'
'                    Res = SaveTransDataW(gRow)
'
'                    If Res = -1 Then
'                        '-- 저장 실패
'                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
'                        SetText vasID, "Failed", gRow, colState
'                    Else
'                        '-- 저장 성공
'                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
'                        SetText vasID, "Trans", gRow, colState
'
'                        SQL = " Update PATRESULT Set " & vbCrLf & _
'                              " sendflag = '2' " & vbCrLf & _
'                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
'                              " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
'                        Res = SendQuery(gLocal, SQL)
'                        If Res = -1 Then
'                            SaveQuery SQL
'                            Exit Sub
'                        End If
'                    End If
'                End If
'
'                'SetText vasID, "Result", gRow, colState
'                strState = ""
'
'            Case "DE"   '## End Result Text
'                strState = ""
'        End Select
'    Next
'
'End Sub

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
    
    SQL = "select resprec, reflow, refhigh from EQPMASTER where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
    Res = GetDBSelectColumn(gLocal, SQL)
    
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
    SetResult = sResult
    
End Function

' asRow1 = Work List
' asRow2 = 결과 List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    
    'sExamDate = Format(dtpToday, "yyyymmdd")
    sExamDate = Trim(GetText(vasID, asRow1, colOrdDate))
    
    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          " WHERE EXAMDATE = '" & Trim(GetText(vasID, asRow1, colOrdDate)) & "' " & vbCrLf & _
          "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
          "   AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
          "   AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "INSERT INTO PATRESULT("
    SQL = SQL & " EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
                " PID,RECENO,PNAME,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
                " EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
    SQL = SQL & "VALUES("
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colOrdDate)) & "', " & vbCrLf
    SQL = SQL & "'" & gEquip & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colInNum)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colGumNum)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPName)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colMachResult)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf
    SQL = SQL & "'0', " & vbCrLf
    SQL = SQL & "'" & gIFUser & "')"
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Private Sub Var_Clear()
    
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

Private Sub txtBarNum_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(txtBarNum) = 12 And IsNumeric(txtBarNum) Then
            Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"), Trim(txtBarNum))
        End If
        vasID.RowHeight(-1) = 12
    End If
    
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
Dim intRow As Integer

    If KeyAscii = 13 Then
        With vasWorkList
            For intRow = .ActiveRow To .DataRowCnt
                '.Row = intRow
                '.Col = colCheckBox
                'If .Value = 1 Then
                    SetText vasWorkList, txtNum, intRow, colPos
    
                    txtNum = Val(txtNum) + 1
                    
                'End If
            Next
        End With
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


Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lblDate.Caption = Trim(GetText(vasID, Row, colOrdDate))
    lsID = Trim(GetText(vasID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPName))
    
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
          SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, SENDFLAG " & vbCrLf
    SQL = SQL & "  FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "'" & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & lsID & "' " & vbCrLf
    SQL = SQL & "   AND EXAMDATE = '" & Trim(GetText(vasID, Row, colOrdDate)) & "' " & vbCrLf
    SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SENDFLAG "
    
    Res = GetDBSelectVas(gLocal, SQL, vasRes)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.MaxRows = vasRes.DataRowCnt
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
'        SQL = " DELETE FROM PATRESULT " & vbCrLf & _
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
'        GetSampleInfoW (iRow)
'
'        lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'        'Local에서 불러오기
'        ClearSpread vasTemp
'
'        '장비코드, 검사코드, 검사명, 결과, 순번
'        SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, SEQNO " & vbCrLf & _
'              "  FROM EQPMASTER " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " ORDER BY SEQNO "
'
'        res = GetDBSelectVas(gLocal, SQL, vasTemp)
'        If res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'        If lsID <> lblChangeBar.Caption Then
'            For i = 1 To 3
'                SQL = "INSERT INTO PATRESULT(EQUIPNO, BARCODE, DISKNO, " & vbCrLf & _
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
'            SQL = " DELETE FROM PATRESULT " & vbCrLf & _
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
'                SQL = "UPDATE PATRESULT "
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
'    res = GetDBSelectColumn(gLocal, SQL)
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
'    SQL = "Select count(*) from QCRESULT " & vbCrLf & _
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
'        SQL = "delete from QCRESULT " & vbCrLf & _
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
'    SQL = "Insert into QCRESULT (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
'          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
'    res = SendQuery(gLocal, SQL)
'    If res = -1 Then
'        'db_RollBack gLocal
'        SaveQuery SQL
'        Exit Function
'    End If
'
'End Function



Private Sub vasRes_KeyPress(KeyAscii As Integer)
    
    With vasRes
        If KeyAscii = 13 And .ActiveCol = 5 Then
            SQL = ""
            SQL = SQL & "UPDATE PATRESULT " & vbCrLf
            SQL = SQL & "   SET RESULT    ='" & Trim(GetText(vasRes, .ActiveRow, colResult)) & "' " & vbCrLf
            SQL = SQL & " WHERE BARCODE   = '" & Trim(lblBarcode(0).Caption) & "' " & vbCrLf
            SQL = SQL & "   AND EXAMDATE  = '" & Trim(lblDate.Caption) & "' "
            SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
            SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRes, .ActiveRow, colExamCode)) & "' " & vbCrLf
            SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRes, .ActiveRow, colEquipCode)) & "' " & vbCrLf
'            SQL = SQL & "   AND PID       = '" & Trim(GetText(vasID, iRow, colPID)) & "' " & vbCrLf
'            SQL = SQL & "   AND DISKNO    = '" & Trim(GetText(vasID, iRow, colRack)) & "' " & vbCrLf
'            SQL = SQL & "   AND POSNO     = '" & Trim(GetText(vasID, iRow, colPos)) & "' " & vbCrLf

            Res = SendQuery(gLocal, SQL)

            If Res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If
            
        End If
    End With
    
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
    
    lsID = Trim(GetText(vasRID, Row, 2))
    lblChangeBar.Caption = lsID
    lblBarcode(1).Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasRID, Row, 5))
    lblPname(1).Caption = Trim(GetText(vasRID, Row, 6))
    lblRrow.Caption = Row
    'Local에서 불러오기
    ClearSpread vasRRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = ""
    SQL = "SELECT EQUIPCODE,EXAMCODE,EXAMNAME,EQUIPRESULT,RESULT,SEQNO,REFFLAG " & vbCrLf & _
          "  FROM PATRESULT " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' " & vbCrLf & _
          " GROUP BY EQUIPCODE,EXAMCODE,EXAMNAME,EQUIPRESULT,RESULT,SEQNO,REFFLAG "
    
    Res = GetDBSelectVas(gLocal, SQL, vasRRes)
    
    If Res = -1 Then
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

'Private Sub vasRID_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim iRow As Long
'    Dim lsID As String
'    Dim lsTime As String
'    Dim lsPid As String
'    Dim i As Integer
'
'    iRow = vasRID.ActiveRow
'
'    If KeyCode = 13 Then
'
'        If GetSampleInfoR(iRow) = -1 Then
'            Exit Sub
'        End If
'
'        lsID = Trim(GetText(vasRID, iRow, colBarcode))
'
'        'Local에서 불러오기
'        ClearSpread vasTemp
'
'        '장비코드, 검사코드, 검사명, 결과, 순번
'        SQL = ""
'        SQL = SQL & "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SEQNO, SENDFLAG " & vbCrLf
'        SQL = SQL & "  FROM PATRESULT " & vbCrLf
'        SQL = SQL & " WHERE EQUIPNO  = '" & gEquip & "' "
'        SQL = SQL & "   AND BARCODE  = '" & lsID & "' " & vbCrLf
'        SQL = SQL & "   AND EXAMDATE = '" & Trim(Format(dtpExamDate.Value, "yyyymmdd")) & "' " & vbCrLf
'        SQL = SQL & " GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, RESULT, SENDFLAG "
'
'        Res = GetDBSelectVas(gLocal, SQL, vasTemp)
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        If lsID <> lblChangeBar.Caption Then
'            For i = 1 To vasRRes.DataRowCnt
'                SQL = ""
'                SQL = SQL & "INSERT INTO PATRESULT("
'                SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
'                            "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
'                            "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
'                SQL = SQL & "VALUES("
'                SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
'                SQL = SQL & "'" & gEquip & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colBarcode)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colDISK)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colPos)) & "', " & vbCrLf
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colPID)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colPName)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colSex)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRID, iRow, colAge)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colEquipCode)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colExamCode)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colSeq)) & "', " & vbCrLf
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colMachResult)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colResult)) & "', "
'                SQL = SQL & "'" & Trim(GetText(vasRRes, i, colExamName)) & "', "
'                SQL = SQL & "'0', "
'                SQL = SQL & "'" & gIFUser & "')"
'
'                Res = SendQuery(gLocal, SQL)
'
'                If Res = -1 Then
'                    SaveQuery SQL
'                    Exit Sub
'                End If
'            Next i
'
'            SQL = ""
'            SQL = SQL & "DELETE FROM PATRESULT " & vbCrLf
'            SQL = SQL & " WHERE EQUIPNO  = '" & gEquip & "' " & vbCrLf
'            SQL = SQL & "   AND BARCODE  = '" & lblChangeBar.Caption & "' " & vbCrLf
'            SQL = SQL & "   AND PID      = '" & lblChangePID.Caption & "' " & vbCrLf
'            SQL = SQL & "   AND DISKNO   = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf
'            SQL = SQL & "   AND POSNO    = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf
'            SQL = SQL & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' "
'
'            Res = SendQuery(gLocal, SQL)
'
'            If Res = -1 Then
'                SaveQuery SQL
'                Exit Sub
'            End If
'
'        ElseIf lsID = lblChangeBar.Caption Then
'            For i = 1 To vasRRes.DataRowCnt
'
'                SQL = ""
'                SQL = SQL & "UPDATE PATRESULT " & vbCrLf
'                SQL = SQL & "   SET RESULT    ='" & Trim(GetText(vasRRes, i, colResult)) & "' " & vbCrLf
'                SQL = SQL & " WHERE BARCODE   = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf
'                SQL = SQL & "   AND EQUIPNO   = '" & gEquip & "' " & vbCrLf
'                SQL = SQL & "   AND EXAMCODE  = '" & Trim(GetText(vasRRes, i, colExamCode)) & "' " & vbCrLf
'                SQL = SQL & "   AND EQUIPCODE = '" & Trim(GetText(vasRRes, i, colEquipCode)) & "' " & vbCrLf
'                SQL = SQL & "   AND PID       = '" & Trim(GetText(vasRID, iRow, colPID)) & "' " & vbCrLf
'                SQL = SQL & "   AND DISKNO    = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf
'                SQL = SQL & "   AND POSNO     = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf
'                SQL = SQL & "   AND EXAMDATE  = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' "
'
'                Res = SendQuery(gLocal, SQL)
'
'                If Res = -1 Then
'                    SaveQuery SQL
'                    Exit Sub
'                End If
'            Next i
'        End If
'    ElseIf KeyCode = vbKeyDelete Then
'        If iRow < 1 Or iRow > vasRID.DataRowCnt Then
'            Exit Sub
'        End If
'
'        lsID = Trim(GetText(vasRID, iRow, colBarcode))
'        lsPid = Trim(GetText(vasRID, iRow, colPID))
'
'        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = ""
'        SQL = SQL & "DELETE FROM PATRESULT " & vbCrLf
'        SQL = SQL & " WHERE EQUIPNO  = '" & gEquip & "' " & vbCrLf
'        SQL = SQL & "   AND BARCODE  = '" & lsID & "' " & vbCrLf
'        SQL = SQL & "   AND PID      = '" & lsPid & "' " & vbCrLf
'        SQL = SQL & "   AND DISKNO   = '" & Trim(GetText(vasRID, iRow, colRack)) & "' " & vbCrLf
'        SQL = SQL & "   AND POSNO    = '" & Trim(GetText(vasRID, iRow, colPos)) & "' " & vbCrLf
'        SQL = SQL & "   AND EXAMDATE = '" & Format(dtpExamDate.Value, "YYYYMMDD") & "' "
'
'        Res = SendQuery(gLocal, SQL)
'
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If
'
'        DeleteRow vasRID, iRow, iRow
'        vasRRes.MaxRows = 0
'
'    End If
'End Sub

Private Sub vasID_KeyPress(KeyAscii As Integer)
Dim intRow As Integer
Dim lngNum As Long

    If KeyAscii = 13 Then
        vasID.Row = vasID.ActiveRow
        vasID.Col = colSeqno
        If Not IsNumeric(vasID.Text) Then
            Exit Sub
        End If
        
        lngNum = vasID.Text
        
        For intRow = vasID.ActiveRow + 1 To vasID.DataRowCnt
            
            lngNum = lngNum + 1
            Call vasID.SetText(colSeqno, intRow, lngNum)
            
        Next
        
        txtNum.Text = lngNum
        
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


Private Sub vasWorkList_Click(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then
        Call SpreadSheetSort(vasWorkList, Col)
    End If
End Sub


Private Sub vasWorkList_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim intRow As Integer
    Dim j  As Integer
    
    If Row = 0 Then
        Exit Sub
    End If

    With vasWorkList
        .Row = Row
        
        vasID.MaxRows = vasID.MaxRows + 1
'        txtNum.Text = txtNum.Text + 1
        
'        .Col = colBarcode
'        SetText vasID, Trim(.Text), vasID.MaxRows, colBarcode
'Const colOrdDate = 2    '검사의뢰일
'Const colBarcode = 3    '챠트번호(=바코드)
'Const colInNum = 4      '내원번호
'Const colGumNum = 5     '검진번호
'Const colSeqno = 6      'SeqNo
'Const colRack = 7       'Rack
'Const colPos = 8        'Pos
'Const colPName = 9      '환자명
'Const colPSex = 10      '성별
        
        SetText vasID, Trim(GetText(vasWorkList, Row, colCheckBox)), vasID.MaxRows, colCheckBox
        SetText vasID, Trim(GetText(vasWorkList, Row, colOrdDate)), vasID.MaxRows, colOrdDate
        SetText vasID, Trim(GetText(vasWorkList, Row, colBarcode)), vasID.MaxRows, colBarcode
        SetText vasID, Trim(GetText(vasWorkList, Row, colInNum)), vasID.MaxRows, colInNum
        SetText vasID, Trim(GetText(vasWorkList, Row, colGumNum)), vasID.MaxRows, colGumNum
        SetText vasID, Trim(GetText(vasWorkList, Row, colRack)), vasID.MaxRows, colRack
        SetText vasID, Trim(GetText(vasWorkList, Row, colPos)), vasID.MaxRows, colPos
        SetText vasID, Trim(GetText(vasWorkList, Row, colPName)), vasID.MaxRows, colPName
        SetText vasID, Trim(GetText(vasWorkList, Row, colPSex)), vasID.MaxRows, colPSex
    
    End With

End Sub


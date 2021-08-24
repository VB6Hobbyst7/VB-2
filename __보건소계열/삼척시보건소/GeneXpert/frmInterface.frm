VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInterface 
   Caption         =   "GeneXpert Interface "
   ClientHeight    =   9765
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   16080
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
   ScaleHeight     =   9765
   ScaleWidth      =   16080
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox Picture1 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   16020
      TabIndex        =   38
      Top             =   0
      Width           =   16080
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   7590
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   " GeneXpert Interface"
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
         TabIndex        =   42
         Top             =   90
         Width           =   2025
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12180
         Picture         =   "frmInterface.frx":14F5
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13380
         Picture         =   "frmInterface.frx":1A7F
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14670
         Picture         =   "frmInterface.frx":2009
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port"
         Height          =   195
         Index           =   0
         Left            =   11640
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Send"
         Height          =   195
         Left            =   12765
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Receive"
         Height          =   195
         Left            =   13800
         TabIndex        =   39
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   3375
      Left            =   15180
      TabIndex        =   31
      Top             =   7230
      Visible         =   0   'False
      Width           =   8625
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   240
         TabIndex        =   32
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
         SpreadDesigner  =   "frmInterface.frx":2593
      End
      Begin FPSpread.vaSpread vasPrintBuf 
         Height          =   1245
         Left            =   240
         TabIndex        =   33
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
         SpreadDesigner  =   "frmInterface.frx":401A
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   6375
      Left            =   15240
      TabIndex        =   9
      Top             =   660
      Visible         =   0   'False
      Width           =   8655
      Begin VB.ComboBox cboChk 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmInterface.frx":4240
         Left            =   1860
         List            =   "frmInterface.frx":424A
         TabIndex        =   79
         Top             =   -30
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton cmdSearch 
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
         Left            =   3720
         TabIndex        =   77
         Top             =   0
         Visible         =   0   'False
         Width           =   975
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
         Left            =   4770
         TabIndex        =   76
         Top             =   0
         Visible         =   0   'False
         Width           =   975
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
         Height          =   345
         Left            =   6690
         TabIndex        =   75
         Text            =   "0"
         Top             =   30
         Visible         =   0   'False
         Width           =   675
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   1290
         Left            =   3090
         TabIndex        =   51
         Top             =   150
         Visible         =   0   'False
         Width           =   1515
         _Version        =   393216
         _ExtentX        =   2672
         _ExtentY        =   2275
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
         MaxCols         =   10
         SpreadDesigner  =   "frmInterface.frx":425A
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   1455
         Left            =   120
         TabIndex        =   23
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
         SpreadDesigner  =   "frmInterface.frx":7E59
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   2235
         Left            =   3780
         TabIndex        =   10
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
         SpreadDesigner  =   "frmInterface.frx":807F
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         Picture         =   "frmInterface.frx":82A5
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   24
         Top             =   5790
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   210
         TabIndex        =   22
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
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   16
         Top             =   4830
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   5100
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   5640
         Value           =   1  '확인
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   4860
         TabIndex        =   11
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
            Left            =   150
            Top             =   210
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
            Handshaking     =   1
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
                  Picture         =   "frmInterface.frx":882F
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":8DC9
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":9363
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":98FD
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":A18F
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":A2E9
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":A443
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1485
         Left            =   120
         TabIndex        =   17
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
         SpreadDesigner  =   "frmInterface.frx":A59D
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   2205
         Left            =   3780
         TabIndex        =   18
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
         SpreadDesigner  =   "frmInterface.frx":A7C3
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1485
         Left            =   120
         TabIndex        =   19
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
         SpreadDesigner  =   "frmInterface.frx":A9E9
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
         Left            =   30
         TabIndex        =   80
         Top             =   450
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label13 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "Seq"
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
         Left            =   6210
         TabIndex        =   78
         Top             =   120
         Visible         =   0   'False
         Width           =   375
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
         TabIndex        =   34
         Top             =   5760
         Width           =   825
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2940
         TabIndex        =   21
         Top             =   5730
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3720
         TabIndex        =   20
         Top             =   5730
         Width           =   705
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   8715
      Left            =   120
      TabIndex        =   1
      Top             =   450
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   15372
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
      TabPicture(0)   =   "frmInterface.frx":AC0F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtTest"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdWorkList"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "이전결과"
      TabPicture(1)   =   "frmInterface.frx":AC2B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "워크조회"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4170
         TabIndex        =   73
         Top             =   540
         Width           =   1485
      End
      Begin VB.Frame Frame3 
         Height          =   8295
         Left            =   -74820
         TabIndex        =   53
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
            TabIndex        =   67
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
            TabIndex        =   66
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
            TabIndex        =   65
            Top             =   210
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   7860
            TabIndex        =   59
            Top             =   630
            Width           =   6675
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   64
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   1
               Left            =   4590
               TabIndex        =   63
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
               TabIndex        =   62
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   1
               Left            =   1995
               TabIndex        =   61
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
               Left            =   510
               TabIndex        =   60
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
            TabIndex        =   58
            Top             =   90
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
            TabIndex        =   57
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   56
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
            TabIndex        =   55
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
            Left            =   13110
            TabIndex        =   54
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   6780
            Left            =   7860
            TabIndex        =   68
            Top             =   1455
            Width           =   6675
            _Version        =   393216
            _ExtentX        =   11774
            _ExtentY        =   11959
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
            SpreadDesigner  =   "frmInterface.frx":AC47
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   69
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
            Height          =   7515
            Left            =   165
            TabIndex        =   70
            Top             =   720
            Width           =   7605
            _Version        =   393216
            _ExtentX        =   13414
            _ExtentY        =   13256
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
            SpreadDesigner  =   "frmInterface.frx":E963
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
            TabIndex        =   72
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
            TabIndex        =   71
            Top             =   390
            Width           =   780
         End
      End
      Begin VB.TextBox txtTest 
         Height          =   375
         Left            =   2730
         TabIndex        =   45
         Top             =   30
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command16 
         Caption         =   "전송테스트"
         Height          =   435
         Left            =   4590
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   8205
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   14625
         Begin VB.TextBox Text1 
            Height          =   5475
            Left            =   8970
            MultiLine       =   -1  'True
            TabIndex        =   74
            Top             =   2370
            Visible         =   0   'False
            Width           =   8655
         End
         Begin VB.TextBox txtUID 
            Appearance      =   0  '평면
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   11310
            TabIndex        =   52
            Top             =   240
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.CheckBox chkWAll 
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   570
            TabIndex        =   47
            Top             =   720
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   7425
            Left            =   120
            TabIndex        =   50
            Top             =   660
            Width           =   7485
            _Version        =   393216
            _ExtentX        =   13203
            _ExtentY        =   13097
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
            SpreadDesigner  =   "frmInterface.frx":F3DC
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasWorkList 
            Height          =   4575
            Left            =   270
            TabIndex        =   43
            Top             =   4770
            Visible         =   0   'False
            Width           =   7185
            _Version        =   393216
            _ExtentX        =   12674
            _ExtentY        =   8070
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            DisplayRowHeaders=   0   'False
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
            GridShowHoriz   =   0   'False
            GridShowVert    =   0   'False
            MaxCols         =   13
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":FF19
            UserResize      =   2
         End
         Begin VB.CommandButton cmdDownload 
            Caption         =   "Down"
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
            Left            =   6000
            TabIndex        =   46
            Top             =   -150
            Visible         =   0   'False
            Width           =   975
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
            Left            =   10800
            TabIndex        =   36
            Top             =   60
            Visible         =   0   'False
            Width           =   735
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
            Left            =   10020
            TabIndex        =   35
            Top             =   60
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdIFClear 
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
            Left            =   13050
            TabIndex        =   8
            Top             =   210
            Width           =   1395
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "&Save"
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
            Left            =   11580
            TabIndex        =   7
            Top             =   210
            Width           =   1395
         End
         Begin VB.Frame Frame6 
            Height          =   585
            Left            =   7710
            TabIndex        =   25
            Top             =   570
            Width           =   6735
            Begin VB.Label Label8 
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
               Left            =   510
               TabIndex        =   30
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   0
               Left            =   1995
               TabIndex        =   29
               Top             =   240
               Width           =   1425
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
               Left            =   3540
               TabIndex        =   28
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   0
               Left            =   4590
               TabIndex        =   27
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   26
               Top             =   720
               Width           =   1155
            End
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   690
            TabIndex        =   5
            Top             =   5160
            Width           =   225
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   6780
            Left            =   7710
            TabIndex        =   6
            Top             =   1305
            Width           =   6675
            _Version        =   393216
            _ExtentX        =   11774
            _ExtentY        =   11959
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
            SpreadDesigner  =   "frmInterface.frx":10A41
         End
         Begin VB.Frame Frame2 
            Caption         =   "Error Log"
            Height          =   1815
            Left            =   8505
            TabIndex        =   3
            Top             =   6720
            Visible         =   0   'False
            Width           =   5970
            Begin VB.TextBox txtErrLog 
               Appearance      =   0  '평면
               Height          =   1455
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   4
               Top             =   240
               Width           =   5775
            End
         End
         Begin MSComCtl2.DTPicker dtpToday 
            Height          =   315
            Left            =   8640
            TabIndex        =   48
            Top             =   240
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
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   315
            Left            =   2580
            TabIndex        =   81
            Top             =   240
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
            Left            =   1110
            TabIndex        =   82
            Top             =   240
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
            Left            =   150
            TabIndex        =   83
            Top             =   300
            Width           =   780
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
            Left            =   7710
            TabIndex        =   49
            Top             =   300
            Width           =   780
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
            Left            =   8955
            TabIndex        =   37
            Top             =   150
            Visible         =   0   'False
            Width           =   780
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   9360
      Width           =   16080
      _ExtentX        =   28363
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4577
            MinWidth        =   4586
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   11466
            MinWidth        =   11466
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   6068
            MinWidth        =   6068
            TextSave        =   "2018-07-06"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "오후 4:34"
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

'vasid, vasrid colum
Const colSpecNo = 0 '미사용
Const colCheckBox = 1
Const colSeqNo = 2
Const colOrdDate = 3
Const colBarcode = 4
Const colRack = 5
Const colPos = 6
Const colPID = 7
Const colPName = 8
Const colSex = 9
Const colAge = 10
Const colOCnt = 11
Const colRCnt = 12
Const colState = 13

'vasres, vasrres colum
Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
Const colMachResult = 4
Const colResult = 5
Const colSeq = 6
Const colFLAG = 7

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
Dim blnLDLCal As Boolean

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
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 1
        Next iRow
    ElseIf chkWAll.Value = 0 Then
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            vasID.Value = 0
        Next iRow
    End If
End Sub

Private Sub cmdDownload_Click()
    Dim intRow As Integer
    Dim j  As Integer
    
    j = 0
    With vasWorkList
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = colCheckBox
            If .Value = 1 Then
                vasID.MaxRows = vasID.MaxRows + 1
                
                .Col = colBarcode
                SetText vasID, txtNum, vasID.MaxRows, colSeqNo
                SetText vasID, Trim(.Text), vasID.MaxRows, colBarcode
                Call GetSampleInfoW(vasID.MaxRows)                                '5,6,7,8
                
                'Call .DeleteRows(intRow, intRow)
                '.MaxRows = .MaxRows - 1
                '.Action = ActionDeleteRow
'                .MaxRows = .MaxRows - 1

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
            SetText vasPrint, Trim(GetText(vasRID, iRow, colPID)), j, 2
            SetText vasPrint, Trim(GetText(vasRID, iRow, colPName)), j, 3
            SetText vasPrint, Trim(GetText(vasRID, iRow, colSex)), j, 4
            
            SQL = " SELECT RESULT " & vbCrLf & _
                  "   FROM PATRESULT " & vbCrLf & _
                  "  WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
                  "    AND EQUIPNO  = '" & gEquip & "' " & vbCrLf & _
                  "    AND BARCODE  = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf & _
                  "    AND PID      = '" & Trim(GetText(vasPrint, iRow, colPID)) & "' " & vbCrLf & _
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
    txtNum = 0
    
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
            
            Res = SaveTransDataW(gRow)
        
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
          SQL = " SELECT '', BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, COUNT(*), COUNT(*), SENDFLAG " & vbCrLf
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
    
    vasRID.RowHeight(-1) = 14

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



Private Sub GetWorkList(ByVal pFrDt As String, ByVal pToDt As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strExamCode As String
    Dim j As Integer
    Dim RS As ADODB.Recordset
    Dim sSpecNo As String
    Dim buff As String
    Dim strTestNm As String
        
    '-- 지누스
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    Dim strMchNum As String
    
    'strMchNum = InputBox("장비번호 입력")
    
'    vasWorkList.MaxRows = 0
    
'    intRow = 0
    
    '-- 로컬 검사코드 찾기
'    Debug.Print gAllExam
    
    '-- 검사대상자 가져오기
                 strRequest = "jobs" + vbTab + "L" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "fr_ymd" + vbTab + pFrDt + vbTab
    strRequest = strRequest & "to_ymd" + vbTab + pToDt + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + "10" + vbTab
    strRequest = strRequest & "smp_no" + vbTab + "%" + vbTab + vbCr
    
    
'ls_service = 'SCC0191A'
'ls_request   = 'jobs' + TAB + 'L' + TAB
'ls_request   += 'fr_ymd' + TAB + '20100301' + TAB
'ls_request   += 'to_ymd' + TAB + '20100331' + TAB
'ls_request   += 'mach' + TAB + '4' + TAB
'ls_request   += 'smp_no' + TAB + '%' + TAB + ENTER
'ls_recv_string = W2ACALL2(ls_service,ls_request, ls_url);
    
    
'                 strRequest = "jobs" + Chr(9) + "L" + Chr(9)
'    strRequest = strRequest & "hos_org_no" + Chr(9) + "31206271" + Chr(9)
'    strRequest = strRequest & "fr_ymd" + Chr(9) + "20100101" + Chr(9)
'    strRequest = strRequest & "to_ymd" + Chr(9) + "20140228" + Chr(9)
'    strRequest = strRequest & "mach" + Chr(9) + strMchNum + Chr(9)
'    strRequest = strRequest & "smp_no" + Chr(9) + "%" + Chr(9) + Chr(13)
    
        
    '"14-0014112-1"
    '"14-0014120-1"
    
    Debug.Print strRequest
'    MsgBox gGINUS_Parm.SVC
'    MsgBox strRequest
'    MsgBox gGINUS_Parm.URL
    
'                     strRequest = "jobs" & vbTab & "L" & vbTab & "hos_org_no" & vbTab & gGINUS_Parm.HCD & vbTab & "fr_ymd" & vbTab & pFrDt & vbTab & "to_ymd" & vbTab & pToDt & vbTab & "mach" & vbTab & strMchNum & vbTab & "smp_no" & vbTab & "%" & vbTab & vbCr
'strRequest = UCase(strRequest)
'    MsgBox strRequest
    
'    strResponse = W2ACALL2(gGINUS_Parm.SVC, strRequest, gGINUS_Parm.URL)
                 
'                 strRequest = "jobs" + vbTab + "Q" + vbTab
'    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
'    strRequest = strRequest & "smp_no" + vbTab + "14-0014120-1" + vbTab
'    strRequest = strRequest & "mach" + Chr(9) + "6" + Chr(9)
'    strResponse = W2ACALL2("SCC0141A", strRequest, "https://121.78.172.70") '-- 바코드로 검사대상 조회
    
'    strResponse = W2ACALL2("SCC0151A", strRequest, "https://121.78.172.70") '-- 모름
'    strResponse = W2ACALL2("SCC0221A", strRequest, "https://121.78.172.70") '-- 채혈대상자
'    strResponse = W2ACALL2("SCC0146A", strRequest, "https://121.78.172.70") '-- 외주

'    strResponse = W2ACALL2("SCC0191A", strRequest, "https://121.78.172.70") '-- 바코드로 검사대상 조회

    
    
    '-- 테스트
    'strResponse = W2ACALL2("HZC0151A", "jobs" & vbTab & "T", "https://121.78.172.70")
    '0         HZC0151A-SUCC                                                                            /CRM/TEST/test.test

'0000009999
'SCC0191A
'W
'Socket Error # 10061
'Connection refused.
    
    Debug.Print strResponse
'          ICD:= Trim(TokenStr(sWork, TAB, 1));     //03  장비코드
'          BCD:= Trim(TokenStr(sWork, TAB, 2));     //10-0001425-1  바코드
'          STA:= Trim(TokenStr(sWork, TAB, 3));     //0: 접수, 1:예비, 2:일부, 3:최종, 4:수정
'          WNO:= Trim(TokenStr(sWork, TAB, 4));     //0
'          ADT:= Trim(TokenStr(sWork, TAB, 5));     //201003220925 접수일시
'          PID:= Trim(TokenStr(sWork, TAB, 6));     //00030617  환자번호
'          PNM:= Trim(TokenStr(sWork, TAB, 7));     //승병순
'          JNO:= Trim(TokenStr(sWork, TAB, 8));     //540324
'          JNO:= JNO + Trim(TokenStr(sWork, TAB, 9));       //JNO + CHARACTER(7) 5140399
'          WAD:= Trim(TokenStr(sWork, TAB, 10));    //5W   병동
'          RUM:= Trim(TokenStr(sWork, TAB, 11));    //507  병실(Room)
'          DPT:= Trim(TokenStr(sWork, TAB, 12));    //FM   접수부서
'          GBN:= Trim(TokenStr(sWork, TAB, 13));    //O : 외래, E : 응급, I : 입원
                  
'                  strResponse = "03" & vbTab & "10-0001425-1" & vbTab & "0" & vbTab & "0" & vbTab & "201003220925" & vbTab & "00030617" & vbTab & "승병순1" & vbTab & "540324" & vbTab & "1010911" & vbTab & "5W" & vbTab & "507" & vbTab & "FM" & vbTab & "O" & vbTab & vbCr
'    strResponse = strResponse & "03" & vbTab & "10-0001425-2" & vbTab & "0" & vbTab & "0" & vbTab & "201003220926" & vbTab & "00030618" & vbTab & "승병순2" & vbTab & "540325" & vbTab & "1010912" & vbTab & "5W" & vbTab & "507" & vbTab & "FM" & vbTab & "I" & vbTab & vbCr
'    strResponse = strResponse & "03" & vbTab & "10-0001425-3" & vbTab & "0" & vbTab & "0" & vbTab & "201003220927" & vbTab & "00030619" & vbTab & "승병순3" & vbTab & "540326" & vbTab & "1010913" & vbTab & "5W" & vbTab & "507" & vbTab & "FM" & vbTab & "E" & vbTab & vbCr
'
'
'                  strResponse = "03" & vbTab & "10-0001425-4" & vbTab & "0" & vbTab & "0" & vbTab & "201003220925" & vbTab & "00030617" & vbTab & "승병순1" & vbTab & "540324" & vbTab & "1010911" & vbTab & "5W" & vbTab & "507" & vbTab & "FM" & vbTab & "O" & vbTab & vbCr
'    strResponse = strResponse & "03" & vbTab & "10-0001425-2" & vbTab & "0" & vbTab & "0" & vbTab & "201003220926" & vbTab & "00030618" & vbTab & "승병순2" & vbTab & "540325" & vbTab & "1010912" & vbTab & "5W" & vbTab & "507" & vbTab & "FM" & vbTab & "I" & vbTab & vbCr
'    strResponse = strResponse & "03" & vbTab & "10-0001425-6" & vbTab & "0" & vbTab & "0" & vbTab & "201003220927" & vbTab & "00030619" & vbTab & "승병순3" & vbTab & "540326" & vbTab & "1010913" & vbTab & "5W" & vbTab & "507" & vbTab & "FM" & vbTab & "E" & vbTab & vbCr
    
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    For i = 0 To UBound(varResponse) - 1
'        If Trim(Mid(varResponse(0), 1, 10)) <> "0" Then
'            MsgBox "조회대상이 없습니다."
'            Exit Sub
'        End If
        If vasWorkList.DataRowCnt = 0 Then
            vasWorkList.MaxRows = 1
            intRow = 1
        
            txtNum = txtNum + 1
            
            SetText vasWorkList, "1", intRow, colCheckBox
            SetText vasWorkList, CStr(txtNum), intRow, colSeqNo
            SetText vasWorkList, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colOrdDate
            SetText vasWorkList, mGetP(varResponse(i), 2, vbTab), intRow, colBarcode
            SetText vasWorkList, mGetP(varResponse(i), 6, vbTab), intRow, colPID
            SetText vasWorkList, mGetP(varResponse(i), 7, vbTab), intRow, colPName
            
            Select Case mGetP(varResponse(i), 13, vbTab)
                Case "O": SetText vasWorkList, "외래", intRow, colRack
                Case "E": SetText vasWorkList, "응급", intRow, colRack
                Case "I": SetText vasWorkList, "입원", intRow, colRack
            End Select
            
            chkWAll.Value = 1
        Else
            '-- Same Check
            intRow = getSameRowNum(Trim(mGetP(varResponse(i), 2, vbTab)))
            If intRow = 0 Then
                vasWorkList.MaxRows = vasWorkList.DataRowCnt + 1
                intRow = vasWorkList.MaxRows
            
                txtNum = txtNum + 1
                
                SetText vasWorkList, "1", intRow, colCheckBox
                SetText vasWorkList, CStr(txtNum), intRow, colSeqNo
                SetText vasWorkList, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colOrdDate
                SetText vasWorkList, mGetP(varResponse(i), 2, vbTab), intRow, colBarcode
                SetText vasWorkList, mGetP(varResponse(i), 6, vbTab), intRow, colPID
                SetText vasWorkList, mGetP(varResponse(i), 7, vbTab), intRow, colPName
                
                Select Case mGetP(varResponse(i), 13, vbTab)
                    Case "O": SetText vasWorkList, "외래", intRow, colRack
                    Case "E": SetText vasWorkList, "응급", intRow, colRack
                    Case "I": SetText vasWorkList, "입원", intRow, colRack
                End Select
            
            End If
        End If
        
'        txtNum = txtNum + 1
'
'        SetText vasWorkList, "1", intRow, colCheckBox
'        SetText vasWorkList, CStr(txtNum), intRow, colSeqNo
'        SetText vasWorkList, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colOrdDate
'        SetText vasWorkList, mGetP(varResponse(i), 2, vbTab), intRow, colBarcode
'        SetText vasWorkList, mGetP(varResponse(i), 6, vbTab), intRow, colPID
'        SetText vasWorkList, mGetP(varResponse(i), 7, vbTab), intRow, colPName
'
'        Select Case mGetP(varResponse(i), 13, vbTab)
'            Case "O": SetText vasWorkList, "외래", intRow, colRack
'            Case "E": SetText vasWorkList, "응급", intRow, colRack
'            Case "I": SetText vasWorkList, "입원", intRow, colRack
'        End Select
    Next
    
    vasWorkList.RowHeight(-1) = 12

End Sub

Private Function getSameRowNum(ByVal strBarNo As String) As Integer
Dim i As Integer

    getSameRowNum = 0
    With vasWorkList
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = colBarcode
            If Trim(.Text) = strBarNo Then
                getSameRowNum = i
                Exit Function
            End If
        Next
    End With
    
End Function

Private Sub cmdSearch_Click()
                
    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    vasID.RowHeight(-1) = 12

End Sub


Private Sub Command1_Click()



End Sub

Private Sub cmdWorkList_Click()

   ' Text1.Text = ""
    
    Call Get_Sample_Info_Work
    
End Sub

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
        
'    MsgBox "2014

'Call Get_OrderList(txtTest.Text)
'
'Exit Sub

    
'    strBuffer = ""
'    strBuffer = strBuffer & "1H|\^&||||||||||P||05" & vbCrLf
'    strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||||||||||3B" & vbCrLf
''
''    Call Get_OrderList("201404240033")
''
''    Exit Sub
'    strBuffer = strBuffer & "3O|1|11208647111|807^00042^3^^SAMPLE^NORMAL|ALL|R|20111205092128|||||X||||||||||||||O|||||38" & vbCrLf
'    strBuffer = strBuffer & "4R|1|^^^321^^0|>100.0|ng/ml|0.000^4.00|>||F|||20111205092406|20111205094226|CF" & vbCrLf
'    strBuffer = strBuffer & "5C|1|I|51^Above measuring range|I04" & vbCrLf
'    strBuffer = strBuffer & "6R|2|^^^391^^0|13.78|ng/ml|^|N||F|||20111205092448|20111205094308|14" & vbCrLf
'    strBuffer = strBuffer & "7L|140" & vbCrLf
'    strBuffer = strBuffer & "" & vbCrLf
'
'    strBuffer = ""
'    strBuffer = strBuffer & "1H|\^&|||iSmart30^iSmart30364^-^1.0.4.1 EX R2||||||||1394-97|2010112514415511" & vbCr
'    strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||||||||||2E" & vbCr
'    strBuffer = strBuffer & "3O|1||3||||||||||||Sample|||||||||||||||80" & vbCr
'    strBuffer = strBuffer & "4R|1|^^^Na+^M|121|mmol/L||N||||||20101125143057|77" & vbCr
'    strBuffer = strBuffer & "5R|2|^^^K+^M|1.9|mmol/L||N|||||||59" & vbCr
'    strBuffer = strBuffer & "6R|3|^^^Cl-^M|73|mmol/L||N|||||||93" & vbCr
'    strBuffer = strBuffer & "7R|4|^^^Hct^M|Out of Range(L)|%||L|||||||38" & vbCr
'    strBuffer = strBuffer & "0L|1|NF693" & vbCr
'    strBuffer = strBuffer & "" & vbCrLf
'
'    strBuffer = "R,NORMAL ,2012-12-21,15:57,2            ,201404240033 ,             ,49,0,000,01,06,Na-PS   ,=,****     mEq/l ,  ,0    ,0    ,        E 8,K-PS    ,=,****     mEq/l ,  ,0.0  ,0.0  ,        E 8,Cl-PS   ,<,-OR      mEq/l ,  ,0    ,0    ,        E 3,GOT-PS  ,=,7        U/l   ,01,0    ,0    , @         ,GOT-PS  ,=,7        U/l   ,01,0    ,0    , @         ,GOT-PS  ,=,7        U/l   ,01,0    ,0    , @         A"
'
'    strBuffer = "R,NORMAL ,2014-04-24,17:22,7            ,1404240099   ,             ,03,0,000,01,15,"
'    strBuffer = strBuffer & "UA-PS   ,=,4.2      mg/dl ,01,4.0  ,7.0  ,           ,"
'    strBuffer = strBuffer & "CRE-PS  ,=,0.6      mg/dl ,01,0.6  ,1.1  ,           ,"
'    strBuffer = strBuffer & "GPT-PS  ,=,23       U/l   ,01,4    ,44   ,           ,"
'    strBuffer = strBuffer & "GOT-PS  ,=,26       U/l   ,01,8    ,38   ,           ,"
'    strBuffer = strBuffer & "TCHO-PS ,=,203      mg/dl ,01,150  ,219  ,             ,"
'    strBuffer = strBuffer & "TG-PS   ,=,107      mg/dl ,01,50   ,149  ,           ,"
'    strBuffer = strBuffer & "CPK-PS  ,=,80       U/l   ,01,40   ,200  ,           ,"
'    strBuffer = strBuffer & "ALP-PS  ,=,169      U/l   ,01,104  ,338  ,           ,"
'    strBuffer = strBuffer & "DBIL-PS ,=,0.1      mg/dl ,01,0.1  ,0.4  , @         ,"
'    strBuffer = strBuffer & "TBIL-PS ,=,0.3      mg/dl ,01,0.1  ,1.2  ,           ,"
'    strBuffer = strBuffer & "ALB-PS  ,=,4.6      g/dl  ,01,3.8  ,5.0  ,           ,"
'    strBuffer = strBuffer & "TP-PS   ,=,7.5      g/dl  ,01,6.7  ,8.3  ,           ,"
'    strBuffer = strBuffer & "BUN-PS  ,=,19.0     mg/dl ,01,8.0  ,23.0 ,           "
'
    strBuffer = "NORMAL 2015-12-1511:041            1512150001  01GPT-P  =22       U/l   1            GOT-P  =21       U/l   1            DBIL-P =0.8      mg/dl 1 H #        "
    
    
    Call comEqp_OnComm
        

    Call Get_OrderList(txtTest.Text)

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
        
    cboChk.ListIndex = 1
    
'    comEqp.CommPort = gSetup.gPort
'    comEqp.RTSEnable = gSetup.gRTSEnable
'    comEqp.DTREnable = gSetup.gDTREnable
'    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
'
'    If comEqp.PortOpen = False Then
'        comEqp.PortOpen = True
'    End If
'
'    If comEqp.PortOpen Then
'        frmInterface.StatusBar1.Panels(2).Text = "연결 되었습니다"
'        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
'        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
''        lblStatus = "작업중.."
'    Else
'        frmInterface.StatusBar1.Panels(2).Text = "연결 되지 않았습니다"
'        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
'        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
''        lblStatus = "작업 대기중.."
'    End If

    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
''    '-- osw 추가
''    For i = 1 To 1
''        If Not Connect_PRServer Then
''            'Cn_Cnt = Cn_Cnt + 1
''            'If Cn_Cnt = 3 Then
''            '    If Not Connect_DRServer Then
''                    MsgBox "연결되지 않았습니다."
''                    cn_Server_Flag = False
''                    Exit Sub
''            '    Else
''            '        cn_Server_Flag = True
''            '    End If
''            'End If
''        Else
''            cn_Server_Flag = True
''        End If
''    Next
    
    
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
    
    Winsock1.LocalPort = CInt(10004)
    Winsock1.Listen
        
    
End Sub



Function GetExamCode() As Integer
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

'Function GetExamCode() As Integer
'    Dim i, j As Long
'
'    ClearSpread vasTemp
'    GetExamCode = -1
'    gAllExam = ""
'    SQL = "Select equipcode, examcode, examname, resprec, seqno " & vbCrLf & _
'          "  From EQPMASTER " & vbCrLf & _
'          " Where equipno = '" & gEquip & "' " & vbCrLf & _
'          " Order by  seqno * 10 "
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
'            gAllExam = "'" & Trim(GetText(vasCode, i, 2)) & "'"
'        Else
'            gAllExam = gAllExam & ",'" & Trim(GetText(vasCode, i, 2)) & "'"
'        End If
'
'        gArrEquip(i, 1) = i
'        For j = 1 To 6
'            gArrEquip(i, j + 1) = Trim(GetText(vasCode, i, j))
'        Next j
'    Next i
'
'    GetExamCode = 1
'End Function
    
    
Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasID
        .MaxCols = colState + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            '.TypeEditCharSet = TypeEditCharSetAlphanumeric
            '.TypeEditCharCase = TypeEditCharCaseSetUpper
            
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            'Call SetText(vasID, gArrEquip(i + 1, 2), 0, colState + (i + 1))
            Call SetText(vasID, gArrEquip(i + 1, 4), 0, colState + (i + 1))
            .ColWidth(colState + (i + 1)) = 10
            
            'cboTest.AddItem gArrEquip(i + 1, 4) & Space(20) & "|" & gArrEquip(i + 1, 3)
        Next
        
'        cboTest.ListIndex = 0
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
            
            SetRawData "[Rx]" & Buffer
            StatusBar1.Panels(2).Text = Buffer
            
            lngBufLen = Len(Buffer)
            
            'Debug.Print Buffer
            
'            For i = 1 To lngBufLen
'                BufChar = Mid$(Buffer, i, 1)
'                Select Case BufChar
'                    Case STX
'                        strBuffer = ""
'                    Case ETX
'                        Call EditRcvDataASTM
'                    Case Else
'                        strBuffer = strBuffer & BufChar
'                End Select
'            Next i
        
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case intPhase
                    Case 1      '## Estabilshment Phase
                        Select Case BufChar
                            Case ENQ
                                intBufCnt = 1
                                Erase strRecvData
                                ReDim Preserve strRecvData(intBufCnt)
                                intPhase = 2
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
                            Case ACK
                                '-- 장비에서 넘어온 시간이 우연히 11:59:59초나 익일에 가까운 시간일 경우
                                '-- 결과 저장시 이전일을 가져올 수 있으므로 날짜를 실시간 업데이트 한다.
                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                dtpToday.Value = Format(strDate, "####-##-##")
                                
                                DoEvents
                                
                                If strState = "Q" Then Call SendOrder
                        
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
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
                            Case vbLf
                            Case vbCr
                                intPhase = 4
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
                        End Select
                    Case 4      '## Termination Phase
                        Select Case BufChar
                            Case STX
                                intPhase = 2
                            Case EOT
                                '-- 장비에서 넘어온 시간이 우연히 11:59:59초나 익일에 가까운 시간일 경우
                                '-- 결과 저장시 이전일을 가져올 수 있으므로 날짜를 실시간 업데이트 한다.
                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                dtpToday.Value = Format(strDate, "####-##-##")

                                DoEvents
                                
                                intPhase = 1
                                
                                Call EditRcvDataASTM
                                
                                'SetRawData "[strState]" & strState

                                If strState = "Q" Then
                                    intSndPhase = 1
                                    intFrameNo = 1
                                    comEqp.Output = ENQ
                                    SetRawData "[Tx]" & ENQ
                                End If
                                
                                
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
        vasID.RowHeight(-1) = 14
    End If
    
    '-- 장비수신정보 표시
    Call SetText(vasID, pBarNo, intRow, colBarcode)             '-- 바코드
    Call SetText(vasID, mResult.RackNo, intRow, colSeqNo)       '-- Seq
    Call SetText(vasID, mResult.TubePos, intRow, colOrdDate)    '-- 검사일자
        
    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    
    Call Get_Sample_Info(intRow)
        
    Call SetText(vasID, "Order", intRow, colState)         '12 진행상태

    gRow = intRow
    
End Sub


Function Get_Sample_Info(ByVal asRow As Long) As Integer
    Dim sGubun As String
    Dim sDate As String
    Dim lsID, lsPID, lsPName, lsAcpDate, lsOrdDate, lsSpcCode, lsExamCode As String
    Dim sBarcode As String
    Dim sRet, sSeg As String
    Dim i, j, k As Integer
    
    Get_Sample_Info = -1
   
    ClearSpread vasOrder
    vasOrder.MaxRows = 50
    'vasOrder.Visible = True
    
    sBarcode = Trim(GetText(vasID, asRow, colBarcode))
    If Len(sBarcode) = 10 Then
        sBarcode = "20" & sBarcode
        vasID.SetText colBarcode, asRow, sBarcode
    End If
    
    sRet = Get_OrderList(sBarcode)
    
'    txtBuff = sRet
    
    sRet = Mid(sRet, InStr(1, sRet, Chr(11)) + 1)
    If InStr(1, sRet, Chr(12)) > 0 Then
        sRet = Left(sRet, InStr(1, sRet, Chr(12)) - 1)
    End If
    
    gOrderExam = ""
    
    i = InStr(1, sRet, Chr(13))
    Do While i > 0
        sSeg = Left(sRet, i - 1)
        sRet = Mid(sRet, i + 1)
        
        Select Case Left(sSeg, InStr(1, sSeg, Chr(124)) - 1)
        Case "MSH"
        Case "PID"
            'PID|||200902240068^이경희^830601^2^20090224^20090224^DefaultDomain^PI
            k = 0
            j = InStr(1, sSeg, Chr(124))
            Do While j > 0
                k = k + 1
                
                If k = 4 Then
                    sSeg = Left(sSeg, j - 1)
                    Exit Do
                End If
                sSeg = Mid(sSeg, j + 1)
                j = InStr(1, sSeg, Chr(124))
            Loop
            k = 0
            j = InStr(1, sSeg, "^")
            Do While j > 0
                k = k + 1
                
                Select Case k
                Case 1
                    lsID = Left(sSeg, j - 1)
                Case 2
                    lsPName = Left(sSeg, j - 1)
                    vasID.SetText colPName, asRow, lsPName
                Case 3
                    lsPID = Left(sSeg, j - 1)
                    'vasID.SetText colPID, asRow, lsPID
                Case 4
                Case 5
                    lsAcpDate = Left(sSeg, j - 1)
                Case 6
                    lsOrdDate = Left(sSeg, j - 1)
                    
                    Exit Do
                End Select
                sSeg = Mid(sSeg, j + 1)
                j = InStr(1, sSeg, "^")
            Loop
            
        Case "PV1"
        Case "OBR"
        Case "OBX"
            'OBX|1|ST|WB2570||||||||R
'            blnLDLCal = False
            k = 0
            j = InStr(1, sSeg, Chr(124))
            Do While j > 0
                k = k + 1
                
                If k = 3 Then
                    lsSpcCode = Left(sSeg, j - 1)
                ElseIf k = 4 Then
                    lsExamCode = Left(sSeg, j - 1)
                    k = vasOrder.DataRowCnt + 1
                    vasOrder.SetText 1, k, lsExamCode
                    vasOrder.SetText 2, k, lsSpcCode
                    If lsExamCode = "WC2430" Then
                        blnLDLCal = True
                    End If
                    gOrderExam = gOrderExam & "'" & Trim(lsExamCode) & "',"
                    Exit Do
                End If
                sSeg = Mid(sSeg, j + 1)
                j = InStr(1, sSeg, Chr(124))
            Loop
        End Select
        
        i = InStr(1, sRet, Chr(13))
    Loop
        
    If gOrderExam <> "" Then
        gOrderExam = Mid(gOrderExam, 1, Len(gOrderExam) - 1)
    End If
        
    Get_Sample_Info = 1


    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
End Function


Function Get_Sample_Info_Work() As Integer
    Dim sGubun As String
    Dim sDate As String
    Dim lsID, lsPID, lsPName, lsAcpDate, lsOrdDate, lsSpcCode, lsExamCode As String
    Dim sBarcode As String
    Dim sRet, sSeg As String
    Dim i, j, k As Integer
    Dim asRow As Long
    Dim strTestCode As String
    
    Get_Sample_Info_Work = -1
   
    ClearSpread vasOrder
    vasOrder.MaxRows = 50
    'vasOrder.Visible = True
    
'    sBarcode = Trim(GetText(vasID, asRow, colBarcode))
'    If Len(sBarcode) = 10 Then
'        sBarcode = "20" & sBarcode
'        vasID.SetText colBarcode, asRow, sBarcode
'    End If
    
    sRet = Get_WorkList
    
    'sRet = Text1.Text
    'Text1.Text = sRet

'gAllExam
    
    
'    sRet = Mid(sRet, InStr(1, sRet, Chr(11)) + 1)
'    If InStr(1, sRet, Chr(12)) > 0 Then
'        sRet = Left(sRet, InStr(1, sRet, Chr(12)) - 1)
'    End If

    gOrderExam = ""
    strTestCode = Replace(gAllExam, "'", "")

    i = InStr(1, sRet, Chr(3))
    Do While i > 0
        sSeg = Left(sRet, i - 1)
        sRet = Mid(sRet, i + 1)

        Debug.Print mGetP(sSeg, 5, Chr(124))
        If InStr(1, strTestCode, mGetP(sSeg, 5, Chr(124))) > 0 Then
            If Format(dtpStartDt.Value, "yyyymmdd") <= mGetP(sSeg, 1, Chr(124)) And Format(dtpStopDt.Value, "yyyymmdd") >= mGetP(sSeg, 2, Chr(124)) Then
                
                
                asRow = asRow + 1
                vasID.MaxRows = asRow
                vasID.SetText colCheckBox, asRow, "1"
                vasID.SetText colPName, asRow, mGetP(sSeg, 7, Chr(124))
                vasID.SetText colPID, asRow, mGetP(sSeg, 6, Chr(124))
                vasID.SetText colBarcode, asRow, mGetP(sSeg, 4, Chr(124)) 'Mid(mGetP(sSeg, 4, Chr(124)), 3)
                vasID.SetText colOrdDate, asRow, mGetP(sSeg, 1, Chr(124))
            End If
            
        End If
        
'        Select Case Left(sSeg, InStr(1, sSeg, Chr(124)) - 1)
'        Case "MSH"
'        Case "PID"
'            'PID|||200902240068^이경희^830601^2^20090224^20090224^DefaultDomain^PI
'            k = 0
'            j = InStr(1, sSeg, Chr(124))
'            Do While j > 0
'                k = k + 1
'
'                If k = 4 Then
'                    sSeg = Left(sSeg, j - 1)
'                    Exit Do
'                End If
'                sSeg = Mid(sSeg, j + 1)
'                j = InStr(1, sSeg, Chr(124))
'            Loop
'            k = 0
'            j = InStr(1, sSeg, "^")
'            Do While j > 0
'                k = k + 1
'
'                Select Case k
'                Case 1
'                    lsID = Left(sSeg, j - 1)
'                Case 2
'                    lsPName = Left(sSeg, j - 1)
'                    vasID.SetText colPName, asRow, lsPName
'                Case 3
'                    lsPID = Left(sSeg, j - 1)
'                    vasID.SetText colPID, asRow, lsPID
'                Case 4
'                Case 5
'                    lsAcpDate = Left(sSeg, j - 1)
'                Case 6
'                    lsOrdDate = Left(sSeg, j - 1)
'
'                    Exit Do
'                End Select
'                sSeg = Mid(sSeg, j + 1)
'                j = InStr(1, sSeg, "^")
'            Loop
'
'        Case "PV1"
'        Case "OBR"
'        Case "OBX"
'            'OBX|1|ST|WB2570||||||||R
''            blnLDLCal = False
'            k = 0
'            j = InStr(1, sSeg, Chr(124))
'            Do While j > 0
'                k = k + 1
'
'                If k = 3 Then
'                    lsSpcCode = Left(sSeg, j - 1)
'                ElseIf k = 4 Then
'                    lsExamCode = Left(sSeg, j - 1)
'                    k = vasOrder.DataRowCnt + 1
'                    vasOrder.SetText 1, k, lsExamCode
'                    vasOrder.SetText 2, k, lsSpcCode
'                    If lsExamCode = "WC2430" Then
'                        blnLDLCal = True
'                    End If
'                    gOrderExam = gOrderExam & "'" & Trim(lsExamCode) & "',"
'                    Exit Do
'                End If
'                sSeg = Mid(sSeg, j + 1)
'                j = InStr(1, sSeg, Chr(124))
'            Loop
'        End Select

        i = InStr(1, sRet, Chr(3))
    Loop

    'If gOrderExam <> "" Then
    '    gOrderExam = Mid(gOrderExam, 1, Len(gOrderExam) - 1)
    'End If

    vasID.RowHeight(-1) = 12
    
    Get_Sample_Info_Work = 1


    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
End Function

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
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If

    '-- 장비수신정보 표시
    Call SetText(vasID, pBarNo, intRow, colBarcode)             '2 Barcode
    'Call SetText(vasID, mResult.RackNo, intRow, colRack)        '3 Rack
    'Call SetText(vasID, mResult.TubePos, intRow, colPos)        '4 Pos
    
    Call vasActiveCell(vasID, intRow, colBarcode)
'
'    '-- 결과스프레드 지우기
    Call ClearSpread(vasRes)
'
'    '-- 검사자 정보 서버테이블 가져와 표시(for 워크리스트)  '5,6,7,8
    Call Get_Sample_Info(intRow)
    
    '-- 현재 Row
    gRow = intRow
    
    

End Sub


'-----------------------------------------------------------------------------'
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
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
    Dim strAbNormal  As String
    
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strtmp      As String
    Dim intIdx      As Integer
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    Dim intCol As Integer
    
    Dim strTC As String
    Dim strTG As String
    Dim strHDL As String
    
    For intCnt = 1 To UBound(strRecvData)
        
        strRcvBuf = strRecvData(intCnt)
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "O"    '## Order
                    strBarNo = Trim(mGetP(strRcvBuf, 3, "|"))
                    
'                    For i = 1 To vasID.DataRowCnt
'                        vasID.Row = i
'                        vasID.Col = colCheckBox
'                        If vasID.Value = "1" Then
'                            gRow = i
'                            strBarNo = GetText(vasID, i, colBarcode)
'                            Exit For
'                        End If
'                    Next i
                
                
                    If Len(strBarNo) = 10 Then
                        strBarNo = "20" & strBarNo
                    End If
                
                    If Len(strBarNo) > 12 Then
                        Exit Sub
                    End If
                    
                    If Not IsNumeric(strBarNo) Then
                        Exit Sub
                    End If
                    
                    If strBarNo = "" Then Exit Sub

                    With mResult
                        .BarNo = strBarNo
                    End With
                                        
                    Call SetPatInfo(strBarNo)
                    
                    strState = "O"
                    
                    '-- 오른쪽 결과화면 초기화
                    vasRes.MaxRows = 0
            
            Case "R"    '## Result
                    'R|30|^Carba-R^^OXA48^^^OXA48^|NEG^|||

                    strIntBase = mGetP(strRcvBuf, 3, "|")
                    strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                    strIntResult = strResult
                    
                    '-- 2018.07.06 음성,양성을 -,+로 변경 요청
'                    If UCase(strResult) = "NOT DETECTED" Then
'                        strResult = "음성"
'                    ElseIf UCase(strResult) = "DETECTED" Then
'                        strResult = "양성"
'                    End If
                    
                    If UCase(strResult) = "NOT DETECTED" Then
                        strResult = "-"
                    ElseIf UCase(strResult) = "DETECTED" Then
                        strResult = "+"
                    End If
                    
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
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
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText vasID, "Result", gRow, colState                 '11 진행상태
                            
                            '-- vasID 에 표시
                            For intCol = colState + 1 To vasID.MaxCols
                                If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                    SetText vasID, strResult, gRow, intCol
                                    Exit For
                                End If
                            Next
    
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                            SetText vasRes, strIntResult, lsResRow, colMachResult     '장비결과
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
                                'strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                '-- Work List
                                SetText vasID, "Result", gRow, colState                 '진행상태
                                    
                                '-- vasID 에 표시
                                For intCol = colState + 1 To vasID.MaxCols
                                    If lsExamCode = gArrEquip(intCol - colState, 3) Then
                                        SetText vasID, strResult, gRow, intCol
                                        Exit For
                                    End If
                                Next
                                
                                '-- 결과 List
                                SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                                SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                                SetText vasRes, strIntResult, lsResRow, colMachResult     '장비결과
                                SetText vasRes, strResult, lsResRow, colResult          '결과
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                '-- 로컬 저장
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                strState = "R"
                            End If
                        End If
                    End If
                
                    
                    vasRes.RowHeight(-1) = 14
                
            Case "C"    '## Comment
            
            Case "L"    '## Terminator
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
                        SetText vasID, "0", gRow, colCheckBox
                        
                        SQL = " Update PATRESULT Set " & vbCrLf & _
                              " sendflag = '2' " & vbCrLf & _
                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
                              " And barcode = '" & Trim(GetText(vasWorkList, gRow, colBarcode)) & "' "
                        
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
Private Sub EditRcvDataAU()
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과
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
    Dim strtmp      As String
    Dim intIdx      As Integer
    
    
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        strType = Mid$(strRcvBuf, 1, 2)
        
        Select Case strType
            '## Order Begin =========================================
            Case "RB"   '## Begin Inquiry Text
            Case "R "    '## Inquiry Order
                strBarNo = Trim(Mid(strRcvBuf, 14, 20))
                strRackNo = Mid(strRcvBuf, 3, 4)
                strTubePos = Mid(strRcvBuf, 7, 2)
                
                With mOrder
                    .BarNo = strBarNo
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .Seq = Mid(strRcvBuf, 9, 5)
                End With
                
                Call GetOrder(strBarNo)
                
            Case "RE"   '## End Inquirty Text
            
            '## Result =========================================
            Case "DB"   '## Begin Result Text
            Case "D "    '## Result
                strBarNo = Trim$(Mid$(strRcvBuf, 14, 10))
                
                With mResult
                    .BarNo = strBarNo
                    .RackNo = Mid(strRcvBuf, 3, 4)
                    .TubePos = Mid(strRcvBuf, 7, 2)
                End With
                
                If strBarNo = "" Then Exit Sub

                strtmp = Mid$(strRcvBuf, 29)
                                
                Call SetPatInfo(strBarNo)
                
                Do While Len(strtmp) >= 11
                    strIntBase = Mid$(strtmp, 2, 2)
                    strResult = Mid$(strtmp, 4, 6)
                    strComm = Mid$(strtmp, 10, 1)
        
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
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
                    strtmp = Mid$(strtmp, 12)
                Loop
                
            
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
                
            Case "DE"   '## End Result Text
                strState = ""
        End Select
    Next

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
    
    sExamDate = Format(dtpToday, "yyyymmdd")

    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
          "  AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
          "  AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "INSERT INTO PATRESULT("
    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
    SQL = SQL & "VALUES("
'    SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
    SQL = SQL & "'" & Trim(Format(dtpToday.Value, "YYYYMMDD")) & "', "
    SQL = SQL & "'" & gEquip & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colRack)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPName)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colSex)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colAge)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colMachResult)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', "
    SQL = SQL & "'0', "
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





Private Sub txtNum_KeyPress(KeyAscii As Integer)
Dim intRow As Integer

    If KeyAscii = 13 Then
        With vasWorkList
            For intRow = .ActiveRow To .DataRowCnt
                '.Row = intRow
                '.Col = colCheckBox
                'If .Value = 1 Then
                    SetText vasWorkList, txtNum, intRow, colSeqNo
    
                    txtNum = Val(txtNum) + 1
                    
                'End If
            Next
        End With
    End If
    
End Sub

'Private Sub vasID_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
'    Dim i As Integer
'
'    For i = BlockRow To BlockRow2
'        vasID.Col = 1
'        vasID.Row = i
'        If vasID.Value = 0 Then
'            vasID.Value = 1
'        Else
'            vasID.Value = 0
'        End If
'    Next i
'End Sub


Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPName))
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, SENDFLAG " & vbCrLf & _
          "FROM PATRESULT " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SENDFLAG "
    SQL = SQL & " ORDER BY SEQNO * 10"
    
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

Private Sub vasID_KeyPress(KeyAscii As Integer)
Dim intRow As Integer
Dim lngNum As Long

    If KeyAscii = 13 Then
        vasID.Row = vasID.ActiveRow
        vasID.Col = colSeqNo
        If Not IsNumeric(vasID.Text) Then
            Exit Sub
        End If
        
        lngNum = vasID.Text
        
        For intRow = vasID.ActiveRow + 1 To vasID.DataRowCnt
            
            lngNum = lngNum + 1
            Call vasID.SetText(colSeqNo, intRow, lngNum)
            
        Next
            
        txtNum.Text = lngNum
        
    End If
    
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
    SQL = SQL & "ORDER BY SEQNO * 10 "
    
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
    
    vasRRes.RowHeight(-1) = 14
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

Private Sub vasRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasRID.ActiveRow
        If lRow < 1 Or lRow > vasRID.DataRowCnt Then Exit Sub
            
        vasRID_Click colBarcode, lRow
    End If
End Sub

Private Sub vasWorkList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    
    If Row < 1 Or Row > vasWorkList.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasWorkList, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasWorkList, Row, colPID))
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasWorkList, Row, colPName))
    'Local에서 불러오기
    ClearSpread vasRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SEQNO, SENDFLAG " & vbCrLf & _
          "FROM PATRESULT " & vbCrLf & _
          "WHERE EQUIPNO = '" & gEquip & "' AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "  AND EXAMDATE = '" & Trim(Format(dtpToday.Value, "yyyymmdd")) & "' " & vbCrLf & _
          "GROUP BY SEQNO, EQUIPCODE, EXAMCODE, EXAMNAME, EQUIPRESULT, RESULT, SENDFLAG "
    
    Res = GetDBSelectVas(gLocal, SQL, vasRes)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.MaxRows = vasRes.DataRowCnt

End Sub

Private Sub vasWorkList_DblClick(ByVal Col As Long, ByVal Row As Long)
'
'    Dim intRow As Integer
'    Dim j  As Integer
'
'    With vasWorkList
'        .Row = Row
'
'        vasID.MaxRows = vasID.MaxRows + 1
'        txtNum.Text = txtNum.Text + 1
'
'        .Col = colBarcode
'        SetText vasID, txtNum, vasID.MaxRows, colSeqNo
'        SetText vasID, Trim(.Text), vasID.MaxRows, colBarcode
'        Call GetSampleInfoW(vasID.MaxRows)
'
''        .Action = ActionDeleteRow
''        .MaxRows = .MaxRows - 1
'    End With

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then
        Winsock1.Close

        Winsock1.Accept requestID
        MsgBox "장비에 접속되었습니다."
    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strText As String
    Dim strtmp As String
    
    Winsock1.GetData strText

    SetRawData "[Rx]" & strText
    StatusBar1.Panels(3).Text = strText

    Call RcvSocketData(strText)
    
End Sub

Public Sub RcvSocketData(ByVal lsData As String)
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    Buffer = lsData

    lngBufLen = Len(Buffer)


    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                        intPhase = 2
                        Winsock1.SendData ACK
                        SetRawData "[Tx]" & ACK
                    Case ACK
                        If strState = "Q" Then Call SendOrder
                
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        Winsock1.SendData ACK
                        SetRawData "[Tx]" & ACK
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
                    Case vbCr
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case vbLf
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
                        Winsock1.SendData ACK
                        SetRawData "[Tx]" & ACK
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        intPhase = 2
                    Case EOT
                        intPhase = 1
                        Call EditRcvDataASTM
                
                End Select
        End Select
    Next i
            
    Exit Sub

ErrHandle:
    Resume Next
End Sub



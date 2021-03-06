VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   " AU680"
   ClientHeight    =   11580
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   19860
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
   ScaleHeight     =   11580
   ScaleWidth      =   19860
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   10245
      Left            =   15240
      TabIndex        =   20
      Top             =   930
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame Frame4 
         Caption         =   "Print"
         Height          =   3375
         Left            =   390
         TabIndex        =   59
         Top             =   6540
         Visible         =   0   'False
         Width           =   8625
         Begin FPSpread.vaSpread vasPrint 
            Height          =   1545
            Left            =   240
            TabIndex        =   60
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
            SpreadDesigner  =   "frmInterface.frx":14F5
         End
         Begin FPSpread.vaSpread vasPrintBuf 
            Height          =   1245
            Left            =   240
            TabIndex        =   61
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
            SpreadDesigner  =   "frmInterface.frx":2F6E
         End
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   960
         Left            =   180
         TabIndex        =   58
         Top             =   1650
         Visible         =   0   'False
         Width           =   3075
         Begin VB.Timer tmrSend 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   2460
            Top             =   210
         End
         Begin VB.Timer tmrReceive 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   1980
            Top             =   210
         End
         Begin MSCommLib.MSComm comEqp 
            Left            =   135
            Top             =   150
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
            Top             =   180
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ImageList imlStatus 
            Left            =   1230
            Top             =   120
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
                  Picture         =   "frmInterface.frx":3186
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":3720
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":3CBA
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":4254
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":4AE6
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":4C40
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":4D9A
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   1545
         Left            =   240
         TabIndex        =   39
         Top             =   3270
         Width           =   3495
         _Version        =   393216
         _ExtentX        =   6165
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
         SpreadDesigner  =   "frmInterface.frx":4EF4
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   150
         TabIndex        =   37
         Top             =   5760
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
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   26
         Top             =   5220
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   4830
         TabIndex        =   25
         Top             =   5160
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
         Left            =   4830
         TabIndex        =   24
         Top             =   5655
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1770
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   23
         Top             =   5220
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
         Left            =   3660
         Style           =   1  '그래픽
         TabIndex        =   22
         Top             =   5190
         Value           =   1  '확인
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   1755
         Left            =   3870
         TabIndex        =   21
         Top             =   3180
         Width           =   4425
         _Version        =   393216
         _ExtentX        =   7805
         _ExtentY        =   3096
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
         SpreadDesigner  =   "frmInterface.frx":510C
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1425
         Left            =   240
         TabIndex        =   27
         Top             =   330
         Width           =   3435
         _Version        =   393216
         _ExtentX        =   6059
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
         SpreadDesigner  =   "frmInterface.frx":5324
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   2175
         Left            =   3750
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   4395
         _Version        =   393216
         _ExtentX        =   7752
         _ExtentY        =   3836
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
         SpreadDesigner  =   "frmInterface.frx":553C
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   240
         TabIndex        =   29
         Top             =   1950
         Width           =   3435
         _Version        =   393216
         _ExtentX        =   6059
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
         SpreadDesigner  =   "frmInterface.frx":5754
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3210
         TabIndex        =   31
         Top             =   5850
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   4050
         TabIndex        =   30
         Top             =   5820
         Width           =   915
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   10455
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   18441
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "frmInterface.frx":596C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "결과확인"
      TabPicture(1)   =   "frmInterface.frx":5988
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   9945
         Left            =   180
         TabIndex        =   12
         Top             =   360
         Width           =   14625
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   8430
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
               BeginProperty Font 
                  Name            =   "돋움체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   4200
               TabIndex        =   36
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label4 
               Caption         =   "환자명 :"
               BeginProperty Font 
                  Name            =   "돋움체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
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
               BeginProperty Font 
                  Name            =   "돋움체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   1
               Left            =   1515
               TabIndex        =   34
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label2 
               Caption         =   "바코드번호 :"
               BeginProperty Font 
                  Name            =   "돋움체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
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
            Caption         =   "엑셀저장"
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
            TabIndex        =   19
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdRSch 
            Caption         =   "결과조회"
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
            Left            =   4170
            TabIndex        =   18
            Top             =   240
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1200
            TabIndex        =   17
            Top             =   270
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   131923968
            CurrentDate     =   40457
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   15
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
            Left            =   5640
            TabIndex        =   14
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdRClear 
            Caption         =   "화면지움"
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
            Left            =   11520
            TabIndex        =   13
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   8535
            Left            =   8430
            TabIndex        =   16
            Top             =   1275
            Width           =   5985
            _Version        =   393216
            _ExtentX        =   10557
            _ExtentY        =   15055
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
            MaxCols         =   6
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":59A4
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   9105
            Left            =   165
            TabIndex        =   49
            Top             =   720
            Width           =   8025
            _Version        =   393216
            _ExtentX        =   14155
            _ExtentY        =   16060
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
            SpreadDesigner  =   "frmInterface.frx":73D5
            UserResize      =   2
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
            TabIndex        =   48
            Top             =   330
            Width           =   780
         End
      End
      Begin VB.Frame Frame1 
         Height          =   9945
         Left            =   -74820
         TabIndex        =   2
         Top             =   360
         Width           =   14625
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "화면지움"
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
            Left            =   11520
            TabIndex        =   11
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "결과전송"
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
            TabIndex        =   10
            Top             =   240
            Width           =   1395
         End
         Begin VB.Frame Frame6 
            Height          =   585
            Left            =   8430
            TabIndex        =   42
            Top             =   630
            Width           =   6015
            Begin VB.Label Label8 
               Caption         =   "바코드번호 :"
               BeginProperty Font 
                  Name            =   "돋움체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               BeginProperty Font 
                  Name            =   "돋움체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   0
               Left            =   1515
               TabIndex        =   46
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label6 
               Caption         =   "환자명 :"
               BeginProperty Font 
                  Name            =   "돋움체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
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
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               BeginProperty Font 
                  Name            =   "돋움체"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   4200
               TabIndex        =   44
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   43
               Top             =   720
               Width           =   1155
            End
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Command16"
            Height          =   645
            Left            =   6240
            TabIndex        =   7
            Top             =   2550
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTest 
            Height          =   675
            Left            =   2100
            TabIndex        =   6
            Top             =   2550
            Visible         =   0   'False
            Width           =   4125
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   5
            Top             =   780
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   9105
            Left            =   165
            TabIndex        =   9
            Top             =   720
            Width           =   8025
            _Version        =   393216
            _ExtentX        =   14155
            _ExtentY        =   16060
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
            SpreadDesigner  =   "frmInterface.frx":7DA2
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8550
            Left            =   8430
            TabIndex        =   8
            Top             =   1275
            Width           =   5985
            _Version        =   393216
            _ExtentX        =   10557
            _ExtentY        =   15081
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
            MaxCols         =   6
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":8637
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
            Left            =   1200
            TabIndex        =   40
            Top             =   270
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   131923968
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
            Left            =   210
            TabIndex        =   41
            Top             =   330
            Width           =   780
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   11175
      Width           =   19860
      _ExtentX        =   35031
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   6350
            MinWidth        =   6350
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12348
            MinWidth        =   12348
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3422
            MinWidth        =   3422
            TextSave        =   "2019-04-09"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "오후 2:07"
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
      Height          =   585
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   15255
      _Version        =   65536
      _ExtentX        =   26908
      _ExtentY        =   1032
      _StockProps     =   15
      Caption         =   "   BECKMAN COULTER AU680 INTERFACE"
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
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8730
         Picture         =   "frmInterface.frx":9FF1
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   56
         Top             =   210
         Width           =   285
      End
      Begin VB.Frame fraRS232 
         Appearance      =   0  '평면
         BackColor       =   &H00AF6523&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   11700
         TabIndex        =   51
         Top             =   90
         Width           =   2985
         Begin VB.Label lblRcv 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "수신"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2010
            TabIndex        =   54
            Top             =   120
            Width           =   420
         End
         Begin VB.Label lblSend 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "송신"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1050
            TabIndex        =   53
            Top             =   120
            Width           =   420
         End
         Begin VB.Label lblPort 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "포트"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   150
            TabIndex        =   52
            Top             =   120
            Width           =   360
         End
         Begin VB.Image imgReceive 
            Height          =   240
            Left            =   2550
            Picture         =   "frmInterface.frx":A57B
            Top             =   90
            Width           =   240
         End
         Begin VB.Image imgSend 
            Height          =   240
            Left            =   1620
            Picture         =   "frmInterface.frx":AB05
            Top             =   90
            Width           =   240
         End
         Begin VB.Image imgPort 
            Height          =   240
            Left            =   690
            Picture         =   "frmInterface.frx":B08F
            Top             =   90
            Width           =   240
         End
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
         Left            =   9120
         TabIndex        =   57
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label5 
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
         TabIndex        =   55
         Top             =   255
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0FFFF&
         Height          =   465
         Left            =   11670
         Shape           =   4  '둥근 사각형
         Top             =   60
         Width           =   3045
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
Const colBarcode = 2
Const colRack = 3
Const colDISK = 3
Const colPos = 4
Const colPID = 5
Const colPName = 6
Const colSex = 7
Const colAge = 8
Const colOCnt = 9
Const colRCnt = 10
Const colState = 11

Const colA1c = 12
Const colIFCC = 13
Const coleAg = 14




'sendflag
'0: Order
'1: Result
'2: Trans

'vasres, vasrres colum
Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
'Const colMachResult = 4
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
            
            SQL = "SELECT RESULT " & vbCrLf & _
                  "FROM PAT_RES " & vbCrLf & _
                  "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
                  "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
                  "  AND BARCODE = '" & Trim(GetText(vasRID, iRow, colBarcode)) & "' " & vbCrLf & _
                  "  AND PID = '" & Trim(GetText(vasPrint, iRow, 3)) & "' " & vbCrLf & _
                  "ORDER BY SEQNO"
            Res = GetDBSelectVas(gLocal, SQL, vasPrintBuf)
            
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
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    lblBarcode(0).Caption = ""
    lblPname(0).Caption = ""
    
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
            '    res = SaveTransDataW_QC(gRow)
            'Else
                Res = SaveTransDataW(gRow)
            'End If
        
            If Res = -1 Then
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
    
    SQL = "SELECT '', BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, COUNT(*), COUNT(*), SENDFLAG " & vbCrLf & _
          "FROM PAT_RES " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND SENDFLAG IN ('0','1', '2') " & vbCrLf & _
          "GROUP BY BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, SENDFLAG"
    Res = GetDBSelectVas(gLocal, SQL, vasRID)
    
    
          '"  AND SENDFLAG IN ('1', '2') "
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For iRow = 1 To vasRID.DataRowCnt
        Select Case Trim(GetText(vasRID, iRow, colState))
        Case "2"
            SetBackColor vasRID, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasRID, "완료", iRow, colState
        Case "0"
            'SetText vasID, "오더", iRow, colState
            SetText vasID, "에러", iRow, colState
        Case "1"
            SetText vasRID, "결과", iRow, colState
        End Select
    Next iRow
    
    vasRID.RowHeight(-1) = 15
    
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
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
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
        
    strBuffer = ""
    strBuffer = strBuffer & "1H|\^&||||||||||P||05" & vbCrLf
    strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||||||||||3B" & vbCrLf
    strBuffer = strBuffer & "3O|1|11208647111|807^00042^3^^SAMPLE^NORMAL|ALL|R|20111205092128|||||X||||||||||||||O|||||38" & vbCrLf
    strBuffer = strBuffer & "4R|1|^^^321^^0|>100.0|ng/ml|0.000^4.00|>||F|||20111205092406|20111205094226|CF" & vbCrLf
    strBuffer = strBuffer & "5C|1|I|51^Above measuring range|I04" & vbCrLf
    strBuffer = strBuffer & "6R|2|^^^391^^0|13.78|ng/ml|^|N||F|||20111205092448|20111205094308|14" & vbCrLf
    strBuffer = strBuffer & "7L|140" & vbCrLf
    strBuffer = strBuffer & "" & vbCrLf
    
    '-- Seq
    'strBuffer = "D 000701 6826      01201206196826    E001    24  102    15  003    46  104  7.59  005  4.96  106  0.13  007    56  108   427  009   144  110  95.4  011    47  112  8.97  013  1.08  114  3.32  015   178  116   147  017    48  118  9.66  019  2.59  120  1.08   23   101   24     2   25     8   26     1   27     3  "
    '-- 바코드
    strBuffer = "D 000201 039903073000126             E01   226H 02    85H 11   9.1  13   141H 18  13.2  21   0.7  24  20.2H "
    
    strBuffer = "DERERBDB"
    strBuffer = "R 003201 0018          1013002058"
    
    strBuffer = "D 003401 0019          1013002058    E      32   1.4  46    26  26  0.81H 01   130  02  3.32L 03  4.29  04   7.3  05   0.5  06   0.1  07   158  09   124H 10   0.7L 11  11.2  12    57  14    39H 15    47H 16    74H 17   259  19   9.1  21   4.7H "
    
    strBuffer = "R 000101 00011013002042"
    
    strBuffer = "D 000101 00011013002042    E012    18  017   129  018    26  "
    
    Call comEqp_OnComm
        

End Sub



Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    If App.PrevInstance Then
        End
    End If
    
On Error GoTo Rst

    Me.Left = 0
    Me.Top = 0
    
    Me.Height = 12465
    Me.Width = 15435
    
    cmdIFClear_Click
    cmdRClear_Click
    lblclear_Click
    
    GetSetup
    
    frmInterface.StatusBar1.Panels(1).Text = "동해동인병원 " '& gUserID
        
    
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
    
    
    GetExamCode
    dtpToday = Date
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -30), "yyyymmdd")
    
    SQL = "DELETE FROM PAT_RES WHERE EXAMDATE < '" & sDate & "'"
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
    
    comEqp.CommPort = gSetup.gPort
    comEqp.RTSEnable = gSetup.gRTSEnable
    comEqp.DTREnable = gSetup.gDTREnable
    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If
    
    If comEqp.PortOpen Then
        StatusBar1.Panels(2).Text = "COM" & comEqp.CommPort & " 포트에 연결 되었습니다"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    Else
        StatusBar1.Panels(2).Text = "통신포트에 연결 되지 않았습니다"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    End If
    
    Exit Sub
    
Rst:
    Frame1.Visible = True
    Frame1.ZOrder 0
    
    If Err.Number = "8002" Then
        If (MsgBox("포트 번호가 잘못되었습니다." & vbNewLine & vbNewLine & "   계속 진행하시겠습니까?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            
            End
        End If
    Else
        MsgBox Err.Number & vbNewLine & Err.Description
    End If
    
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
    
'''                    Case 0      'Message Header
'''                        MHead = "1H|\^&||||||||||P"
'''                        brCom.Output = STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf
'''                        SendCount = SendCount + 1
'''                        Debug.Print "[HOST] " & STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf
'''                        Print #1, "[HOST] " & STX & MHead & vbCr & ETX & MakeCS(MHead) & vbCr & vbLf & Chr(13) + Chr(10);
'''                        MHead = ""
'''                    Case 1      'patient information
'''                        Pinfo = "2P|1||" & PatientID & "|||||||||||||||||||||||||||||||"
'''                        brCom.Output = STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf
'''                        SendCount = SendCount + 1
'''                        Debug.Print "[HOST] " & STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf
'''                        Print #1, "[HOST] " & STX & Pinfo & vbCr & ETX & MakeCS(Pinfo) & vbCr & vbLf & Chr(13) + Chr(10);
'''                        Pinfo = ""
''''                        PatientID = ""
'''                    Case 2      'Test Order
'''                        SendCount = SendCount + 1
'''                        Call OrderingTheDataElecsys(brCom, com_sTemp, brSpread, brChannel, brItemdeci)

'''                        Orderoutput = "3O" & "|1|" & PatientID & "|" & PatientSeq & "|" & OutPutData & "|R|" & Format(Now, "YYYYMMDDHHMMSS") & "|||||N||||||||||||||Q"
'''                        OutPutData = STX & Orderoutput & vbCr & ETX & MakeCS(Orderoutput) & vbCr & vbLf

'''                    Case 3      'Message Terminator
'''                        SendCount = SendCount + 1
'''                        brCom.Output = STX & "4L|1|F" & vbCr & ETX & "FF" & vbCr & vbLf
'''                        Debug.Print "[HOST] " & STX & "4L|1|F" & vbCr & ETX & "FF" & vbCr & vbLf
'''                        Print #1, "[HOST] " & STX & "4L|1|F" & vbCr & ETX & "FF" & vbCr & vbLf & Chr(13) + Chr(10);
'''                    Case Else
'''                        brCom.Output = EOT
'''                        Debug.Print "[HOST] " & EOT
'''                        Print #1, "[HOST] " & EOT & Chr(13) + Chr(10);
'''                        SendCount = 0
'''                        Flag_HQL = ""
    
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
            comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
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

Private Sub comEqp_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long
    Dim strDate As String

    
    Select Case comEqp.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long

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
            lngBufLen = Len(Buffer)
            
            Debug.Print Buffer

            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case BufChar
                    Case STX
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETB
                    Case ETX
                        '-- 장비에서 넘어온 시간이 우연히 11:59:59초나 익일에 가까운 시간일 경우
                        '-- 결과 저장시 이전일을 가져올 수 있으므로 날짜를 실시간 업데이트 한다.
                        'strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                        dtpToday.Value = Now 'Format(strDate, "####-##-##")
                        
                        'DoEvents
                    
                        Call EditRcvData
                    Case Else
                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
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
    
    Call SetText(vasID, pBarNo, intRow, colBarcode)         '2
    Call SetText(vasID, mOrder.RackNo, intRow, colRack)     '3
    Call SetText(vasID, mOrder.TubePos, intRow, colPos)     '4
    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    
    Call GetSampleInfoW(intRow)                            '5,6,7,8
    
    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)
    
    '-- 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.
    '-- intRow 추가
    strItems = GetGetEquipExamCode_AU480(gEquip, pBarNo, intRow)

    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = ""
        'S 003401 0019          1013001918    E
        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
        'Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
        
        SetRawData "[Tx]" & STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
        
        'S 003401 0019          1013001918    E      01020304050607091011121415161719212632
        '동해동인 바코드 10자리
        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
        'Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
        
        SetRawData "[Tx]" & STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
        
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
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    Call SetText(vasID, pBarNo, intRow, colBarcode)             '2 Barcode
    Call SetText(vasID, mResult.RackNo, intRow, colRack)        '3 Rack
    Call SetText(vasID, mResult.TubePos, intRow, colPos)        '4 Pos
    Call vasActiveCell(vasID, intRow, colBarcode)
    
    Call ClearSpread(vasRes)
    
    Call GetSampleInfoW(intRow)                                '5,6,7,8
    
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
'    Dim blnPSA       As Boolean
'    Dim blnfPSA      As Boolean
'    Dim strPSA       As String
'    Dim strfPSA      As String
    Dim strTmp      As String
    Dim intIdx      As Integer
    
    Dim strGFR      As String
    Dim strSEX      As String
    Dim strCREA     As String
    Dim strAGE      As String
    
'    blnPSA = False
'    blnfPSA = False
    
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        SetRawData "[Rx]" & strRcvBuf
        
        strType = Mid$(strRcvBuf, 1, 2)
        
        Select Case strType
            Case "R "    '## Inquiry Order
                'R 000101 0001                1608270009
                strBarno = Trim(Mid(strRcvBuf, 14, 20))
                strRackNo = Mid(strRcvBuf, 3, 4)
                strTubePos = Mid(strRcvBuf, 7, 2)
                
                mOrder.BarNo = strBarno
                mOrder.RackNo = strRackNo
                mOrder.TubePos = strTubePos
                mOrder.Seq = Mid(strRcvBuf, 9, 5)
                'R 003201 0018          1013001917
                'S 003201 0018          1013001917    E      13
                               
                Call GetOrder(strBarno)
                
                '===========================================================================

            Case "D "    '## Result
                'D 000101 00011019004608    0001   7.8  002   4.6  003  0.88  005    26  006    20  007    54  008   176  009    19  010  17.7  011  0.80  013   167  014    73  015    44  017    88  018  10.1  019   3.3  020   6.3  097   139  098   4.5  099   104

                '동해동인 바코드 10자리
                strBarno = Trim$(Mid$(strRcvBuf, 14, 10))
                mResult.BarNo = strBarno
                mResult.RackNo = Mid(strRcvBuf, 3, 4)
                mResult.TubePos = Mid(strRcvBuf, 7, 2)
                
                If strBarno = "" Then Exit Sub

                'intIdx = InStr(strRcvBuf, "E")
                'intIdx = InStr(strRcvBuf, "E")
                strTmp = Mid$(strRcvBuf, 29)
                                
                Call SetPatInfo(strBarno)
                
                strSEX = GetText(vasID, gRow, colSex)
                strAGE = GetText(vasID, gRow, colAge)
                strGFR = ""
                strCREA = ""
                
                Do While Len(strTmp) >= 11
                    strIntBase = Mid$(strTmp, 1, 3)
                    strResult = Trim(Mid$(strTmp, 4, 6))
                    strComm = Mid$(strTmp, 10, 1)
                    
                    If IsNumeric(strResult) Then
                        Select Case strIntBase
                            Case "109"   'H-RPR VDRL 정밀   '47
                                                Select Case Val(strResult)
                                                    Case "0.0":         strResult = "Neg(0.0)"
                                                    Case Is <= 0:       strResult = "Neg(0.0)"
                                                    Case 0.1 To 0.9:    strResult = "Neg(" & strResult & ")"
                                                    Case 1 To 20:       strResult = "Pos(" & strResult & ")"
                                                    Case Is > 20:       strResult = "Pos(>20)"
                                                End Select
                            Case "005"    'AST 08
                                                If CCur(strResult) < 3 Then strResult = "< 3"
                            Case "006"    'ALT 09
                                                If CCur(strResult) < 3 Then strResult = "< 3"
                            Case "004"    'D-Bil 06
                                                If CCur(strResult) < 0.1 Then strResult = "< 0.1"
                            Case "011"    'CREA 14
                                    'GFR 계산
                                    strGFR = ""
                                    strCREA = strResult
                                    
                                    If CCur(strResult) > 0 Then
                                        '18세 이상만 적용
                                        If IsNumeric(strCREA) And strAGE > 18 Then
                                            If strSEX = "M" Then
                                                strGFR = 186 * (strCREA ^ -1.154) * (strAGE ^ -0.203)
                                            ElseIf strSEX = "F" Then
                                                strGFR = 186 * (strCREA ^ -1.154) * (strAGE ^ -0.203) * 0.742
                                            End If
                                            
                                            If strGFR <> "" Then
                                                strGFR = Format(strGFR, "##0.00")
                                                If strGFR <= 120 Then
                                                    strGFR = Round(strGFR, 2)
                                                ElseIf strGFR > 120 Then
                                                    strGFR = "> 120"
                                                End If
                                            End If
                                        End If
                                    Else
                                        strGFR = "Error"
                                    End If
                                    
                                    If CCur(strResult) < 0.2 Then strResult = "< 0.2"
                                    
                                    
                            Case "009"    'GGT  12
                                                If CCur(strResult) < 3 Then strResult = "< 3"
                            Case "020"    'UA   25
                                                If CCur(strResult) < 1.5 Then strResult = "< 1.5"
                            Case "019"    'P    24
                                                If CCur(strResult) < 1 Then strResult = "< 1"
                            Case "012"    'Lipase   45
                                                If CCur(strResult) < 3 Then strResult = "< 3"
'                            Case "32"    'Fe
'                                                If CCur(strResult) < 10 Then strResult = "< 10"
                            Case "025"    'hs-CRP   46
                                                If CCur(strResult) < 0.08 Then strResult = "< 0.08"
                            
                        End Select
                    End If
                    
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQUIPEXAM"
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
                            SetText vasID, strResult, gRow, colA1c                  '결과
                            SetText vasID, strComm, gRow, colA1c + 1                'Flag
                            SetText vasID, "Result", gRow, colState                 '진행상태
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
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
                            SQL = SQL & "  From equipexam"
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
                                SetText vasID, strResult, gRow, colA1c                  '결과
                                SetText vasID, strComm, gRow, colA1c + 1                'Flag
                                SetText vasID, "Result", gRow, colState                 '진행상태
                                '-- 결과 List
                                SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                                SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                                SetText vasRes, strResult, lsResRow, colResult          '결과
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                '-- 로컬 저장
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                strState = ""
                            End If
                        End If
                    End If
                    
                    strTmp = Mid$(strTmp, 12)
                Loop
                strState = "R"
                
                '-- GFR 저장
                If strGFR <> "" Then
                    strIntBase = "088"
                    strResult = strGFR
                    
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQUIPEXAM"
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
                        SetText vasID, strResult, gRow, colA1c                  '결과
                        SetText vasID, strComm, gRow, colA1c + 1                'Flag
                        SetText vasID, "Result", gRow, colState                 '진행상태
                        '-- 결과 List
                        SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                        SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                        SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
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
                        SQL = SQL & "  From equipexam"
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
                            SetText vasID, strResult, gRow, colA1c                  '결과
                            SetText vasID, strComm, gRow, colA1c + 1                'Flag
                            SetText vasID, "Result", gRow, colState                 '진행상태
                            '-- 결과 List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '검사코드
                            SetText vasRes, lsExamName, lsResRow, colExamName       '검사명
                            SetText vasRes, strResult, lsResRow, colResult          '결과
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '순번
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- 로컬 저장
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            strState = ""
                        End If
                    End If
                End If
                
                vasID.RowHeight(-1) = 12
                vasRes.RowHeight(-1) = 12
                
                If MnTransAuto.Checked = True Then
                    
                    Res = SaveTransDataW(gRow)
                    
                    If Res = -1 Then
                        '-- 저장 실패
                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                        SetText vasID, "Failed", gRow, colState
                    Else
                        '-- 저장 성공
                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                        SetText vasID, "Trans", gRow, colState
                        
                        SQL = " Update pat_res Set " & vbCrLf & _
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
            
                SetText vasID, "Result", gRow, colState
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
    
    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
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
    SQL = SQL & "DELETE FROM PAT_RES " & vbCrLf
    SQL = SQL & " WHERE EXAMDATE    = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf
    SQL = SQL & "   AND EQUIPNO     = '" & gEquip & "' " & vbCrLf
    SQL = SQL & "   AND BARCODE     = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf
    SQL = SQL & "   AND EQUIPCODE   = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf
    'SQL = SQL & "   AND EXAMCODE    = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "INSERT INTO PAT_RES("
    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
    SQL = SQL & "VALUES("
    SQL = SQL & "'" & Trim(Format(dtpToday.Value, "YYYYMMDD")) & "', "
    SQL = SQL & "'" & gEquip & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colDISK)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPName)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colSex)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colAge)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
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
    
    If lblUser.Caption <> "" Then
        gUserID = lblUser.Caption
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
    
    vasRes.MaxRows = 0
    
    lsID = Trim(GetText(vasID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPName))
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
    
    Res = GetDBSelectVas(gLocal, SQL, vasRes)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If

    vasRes.MaxRows = vasRes.DataRowCnt
    vasRes.RowHeight(-1) = 12

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
'        GetSampleInfoW (iRow)
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
'        res = GetDBSelectVas(gLocal, SQL, vasTemp)
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

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sRow        As Long
    Dim sCol        As Long
    Dim strNewBarNo As String
    Dim intRow      As Integer
    Dim strSeq      As String
    Dim sExamDate   As String
    
    If lblBarcode(0).Caption = "" Then
        Exit Sub
    End If
    
    sExamDate = Format(dtpToday, "yyyymmdd")
    
    sRow = vasID.ActiveRow
    sCol = vasID.ActiveCol
    
    If sRow = 0 Then
        Exit Sub
    End If
    
    strNewBarNo = GetText(vasID, sRow, sCol)
    
    If KeyCode = vbKeyReturn Then
        If colBarcode = sCol Then
            If GetSampleInfoW(sRow) = -1 Then
                MsgBox "입력한 바코드에서 환자정보를 찾지 못했습니다." & vbNewLine & " 바코드 번호를 확인하세요", vbOKOnly + vbCritical, Me.Caption
            Else
                '정보수정
                SQL = ""
                SQL = SQL & "UPDATE PAT_RES SET "
                SQL = SQL & "  BARCODE = '" & strNewBarNo & "'" & vbCr
                SQL = SQL & " ,PID     = '" & Trim(GetText(vasID, sRow, colPID)) & "'" & vbCr
                SQL = SQL & " ,PNAME   = '" & Trim(GetText(vasID, sRow, colPName)) & "'" & vbCr
                SQL = SQL & " ,PSEX    = '" & Trim(GetText(vasID, sRow, colSex)) & "'" & vbCr
                SQL = SQL & " ,PAGE    = '" & Trim(GetText(vasID, sRow, colAge)) & "'" & vbCr
                SQL = SQL & " WHERE Mid(EXAMDATE,1,8) = '" & sExamDate & "'" & vbCr
                SQL = SQL & "   AND EQUIPNO  = '" & gEquip & "'" & vbCr
                SQL = SQL & "   AND BARCODE  = '" & lblBarcode(0).Caption & "'" & vbCr

                Res = SendQuery(gLocal, SQL)
                
                If Res = -1 Then
                    SaveQuery SQL
                End If
            End If
        End If
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
    Dim strCREA As String
    Dim strAGE As String
    Dim strSEX As String
    Dim strGFR As String
    
    If Row < 1 Or Row > vasRID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasRID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblBarcode(1).Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasRID, Row, colPID))
    lblPname(1).Caption = Trim(GetText(vasRID, Row, colPName))
    lblRrow.Caption = Row
    'Local에서 불러오기
    ClearSpread vasRRes
    
    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = ""
    SQL = "SELECT EQUIPCODE,EXAMCODE,EXAMNAME,EQUIPRESULT,RESULT,SEQNO,REFFLAG " & vbCrLf & _
          "  FROM PAT_RES " & vbCrLf & _
          " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "   AND BARCODE = '" & lsID & "' " & vbCrLf & _
          "   AND DISKNO = '" & Trim(GetText(vasRID, Row, colRack)) & "' " & vbCrLf & _
          "   AND POSNO = '" & Trim(GetText(vasRID, Row, colPos)) & "' " & vbCrLf & _
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
        
        'CREA : 14
'        If Trim(GetText(vasRRes, i, colEquipCode)) = "14" Then
'            strSEX = Trim(GetText(vasRID, Row, colSex))
'            strAGE = Trim(GetText(vasRID, Row, colAge))
'            strCREA = Trim(GetText(vasRRes, i, colResult))
'            'MsgBox strSEX
'            'MsgBox strAGE
'            If IsNumeric(strCREA) Then
'                If strSEX = "M" Then
'                    strGFR = 186 * (strCREA ^ -1.154) * (strAGE ^ -0.203)
'                ElseIf strSEX = "F" Then
'                    strGFR = 186 * (strCREA ^ -1.154) * (strAGE ^ -0.203) * 0.742
'                End If
'            'MsgBox strGFR
'
'                If strGFR <> "" Then
'                    strGFR = Format(strGFR, "##0.00")
'                    If strGFR <= 90 Then
'                        strGFR = Round(strGFR, 2)
'                    ElseIf strGFR > 90 Then
'                        strGFR = ">90"
'                    End If
'                End If
'             '               MsgBox strGFR
'
'                vasRRes.MaxRows = vasRRes.MaxRows + 1
'                Call SetText(vasRRes, "88", vasRRes.MaxRows, colEquipCode)
'                Call SetText(vasRRes, "GFR", vasRRes.MaxRows, colExamName)
'                Call SetText(vasRRes, strGFR, vasRRes.MaxRows, colResult)
'            End If
'        End If
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
'        SQL = SQL & "  FROM PAT_RES " & vbCrLf
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
'                SQL = SQL & "INSERT INTO PAT_RES("
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
'            SQL = SQL & "DELETE FROM PAT_RES " & vbCrLf
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
'                SQL = SQL & "UPDATE PAT_RES " & vbCrLf
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
'        SQL = SQL & "DELETE FROM PAT_RES " & vbCrLf
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

'Private Sub vasRRes_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = 13 Then: vasRID_KeyDown KeyCode, 0
'End Sub

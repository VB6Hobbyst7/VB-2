VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   0  '없음
   Caption         =   " VISTA 500  Interface "
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
      Height          =   2655
      Left            =   420
      TabIndex        =   28
      Top             =   7050
      Visible         =   0   'False
      Width           =   8175
      Begin FPSpread.vaSpread vasCode 
         Height          =   915
         Left            =   7200
         TabIndex        =   47
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
         TabIndex        =   45
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
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   1320
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
         SpreadDesigner  =   "frmInterface.frx":07A5
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
         SpreadDesigner  =   "frmInterface.frx":09BD
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
         SpreadDesigner  =   "frmInterface.frx":0BD5
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   1440
         TabIndex        =   48
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
         SpreadDesigner  =   "frmInterface.frx":0DED
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   4860
         TabIndex        =   39
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   5700
         TabIndex        =   38
         Top             =   1410
         Width           =   915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   2085
      Left            =   15165
      TabIndex        =   25
      Top             =   3915
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
         SpreadDesigner  =   "frmInterface.frx":1005
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
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   8775
         Left            =   -74820
         TabIndex        =   16
         Top             =   360
         Width           =   14625
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   8460
            TabIndex        =   40
            Top             =   630
            Width           =   6015
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   46
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Left            =   4200
               TabIndex        =   44
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
               TabIndex        =   43
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Left            =   1605
               TabIndex        =   42
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
               TabIndex        =   41
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
            TabIndex        =   23
            Top             =   240
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   240
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
            Format          =   94633984
            CurrentDate     =   40457
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   780
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   7815
            Left            =   240
            TabIndex        =   20
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
            SpreadDesigner  =   "frmInterface.frx":2CCE
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   7275
            Left            =   8460
            TabIndex        =   21
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
      End
      Begin VB.Frame Frame1 
         Height          =   8775
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   14625
         Begin VB.CommandButton Command16 
            Caption         =   "Command16"
            Height          =   435
            Left            =   6060
            TabIndex        =   11
            Top             =   4950
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtTest 
            Height          =   675
            Left            =   1680
            TabIndex        =   10
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   270
            Width           =   1395
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   660
            TabIndex        =   9
            Top             =   780
            Width           =   225
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   7815
            Left            =   135
            TabIndex        =   13
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
         Begin FPSpread.vaSpread vasRes 
            Height          =   7800
            Left            =   8490
            TabIndex        =   12
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
            SpreadDesigner  =   "frmInterface.frx":7EB3
         End
         Begin VB.Frame Frame2 
            Caption         =   "Error Log"
            Height          =   1695
            Left            =   8490
            TabIndex        =   8
            Top             =   6840
            Visible         =   0   'False
            Width           =   5970
            Begin VB.ListBox List1 
               BeginProperty Font 
                  Name            =   "돋움"
                  Size            =   8.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   1380
               Left            =   60
               TabIndex        =   49
               Top             =   210
               Width           =   5805
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
            TextSave        =   "2012-02-13"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오전 9:16"
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
      Caption         =   "     VISTA 500  INTERFACE"
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
         Left            =   4785
         Picture         =   "frmInterface.frx":BBFE
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
         Format          =   94633984
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
strBuffer = strBuffer & "1H|\^&|||serial HOST|||||||P|DB 100102|20110811133315|" & vbCrLf
strBuffer = strBuffer & "P|1||||^^^|||U|^||||^^^^|||||^|||||\||^|||||^|||||" & vbCrLf
strBuffer = strBuffer & "O|1|11169701511||^^^1802^2202^RF|R||||||N|||20110811122749|S^|^^^^||1|1|||20110811133315|||F||||||" & vbCrLf
strBuffer = strBuffer & "C|1|I|G001^JP000371-1|G|" & vbCrLf
strBuffer = strBuffer & "R|1|^^^180225" & vbCrLf
strBuffer = strBuffer & "2^2202^RF|18.3|IU/mL|0-15|||F||Siemens|20110811133315|20110811131301|VistaDim1|" & vbCrLf
strBuffer = strBuffer & "C|1|I|R001^10341MA|R|" & vbCrLf
strBuffer = strBuffer & "C|2|I|R002^20111005000000|R|" & vbCrLf
strBuffer = strBuffer & "C|3|I|R003^000071|R|" & vbCrLf
strBuffer = strBuffer & "C|4|I|E511^Above Reference Range|E|" & vbCrLf
strBuffer = strBuffer & "L|1|N|" & vbCrLf
strBuffer = strBuffer & "87" & vbCrLf
strBuffer = strBuffer & "" & vbCrLf
    
    
strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||serial HOST|||||HOST||P|DB 100102|20110812145508|" & vbCr
strBuffer = strBuffer & "Q|1|^1588800040||||||||||O|" & vbCr
strBuffer = strBuffer & "L|1|N|" & vbCr
strBuffer = strBuffer & "03" & vbCr

strBuffer = strBuffer & "" & vbCr
    
    
strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||serial HOST|||||||P|DB 100102|20111013101146|" & vbCrLf
strBuffer = strBuffer & "P|1||||^^^|||U|^||||^^^^|||||^|||||\||^|||||^|||||" & vbCrLf
strBuffer = strBuffer & "O|1|9905700111||^^^0126^2202^ASL\^^^1802^2202^RF|R||||||N|||20111012091129|S^|^^^^||1|2|||20111013101146|||F||||||" & vbCrLf
strBuffer = strBuffer & "C|1|I|G001^PP000374-AA" & vbCrLf

strBuffer = strBuffer & "21|G|" & vbCrLf
strBuffer = strBuffer & "R|1|^^^0126^2202^ASL|||-|||F||Siemens|20111013101146|20111013101145|VistaDim1|" & vbCrLf
strBuffer = strBuffer & "C|1|I|E133^Manually Cancelled Test|E|" & vbCrLf
strBuffer = strBuffer & "R|2|^^^1802^2202^RF|||-|||F||Siemens|20111013101146|20111013101145|VistaDim1|" & vbCrLf
strBuffer = strBuffer & "C|1|I|E133^Manually Cancelled Test|E|" & vbCrLf
strBuffer = strBuffer & "L|31" & vbCrLf

strBuffer = strBuffer & "31|N|" & vbCrLf
strBuffer = strBuffer & "BA" & vbCrLf

strBuffer = strBuffer & "" & vbCrLf
    
strBuffer = ""
strBuffer = strBuffer & ""
strBuffer = strBuffer & "1H|\^&|||serial HOST|||||||P|DB 100102|20111025090507|" & vbCrLf
strBuffer = strBuffer & "P|1||||^^^|||U|^||||^^^^|||||^|||||\||^|||||^|||||" & vbCrLf
strBuffer = strBuffer & "O|1|1120047241||^^^0802^2202^BHCG|R||||||N|||20111025085155|S^|^^^^||1|1|||20111025090507|||F||||||" & vbCrLf
strBuffer = strBuffer & "C|1|I|G001^PT000374-1|G|" & vbCrLf
strBuffer = strBuffer & "R|1|^^^0803E" & vbCrLf

strBuffer = strBuffer & "22^2202^BHCG|||1-3|||F||Siemens|20111025090507|20111025094857|VistaDim1|" & vbCrLf
strBuffer = strBuffer & "C|1|I|R001^11103AA|R|" & vbCrLf
strBuffer = strBuffer & "C|2|I|R002^20120412000000|R|" & vbCrLf
strBuffer = strBuffer & "C|3|I|R003^000844|R|" & vbCrLf
strBuffer = strBuffer & "C|4|I|E541^Above Cutoff Value|E|" & vbCrLf
strBuffer = strBuffer & "L|1|N|" & vbCrLf
strBuffer = strBuffer & "B2" & vbCrLf

strBuffer = strBuffer & "" & vbCrLf
    
strBuffer = ""
strBuffer = strBuffer & ""
strBuffer = strBuffer & "1H|\^&|||serial HOST|||||||P|DB 100102|20111026133929|" & vbCrLf
strBuffer = strBuffer & "P|1||||^^^|||U|^||||^^^^|||||^|||||\||^|||||^|||||" & vbCrLf
strBuffer = strBuffer & "O|1|1120035501||^^^0802^2202^BHCG|R||||||N|||20111025085155|S^|^^^^||1|1|||20111026133929|||F||||||" & vbCrLf
strBuffer = strBuffer & "C|1|I|G001^PF000374-1|G|" & vbCrLf
strBuffer = strBuffer & "R|1|^^^0803E" & vbCrLf

strBuffer = strBuffer & "22^2202^BHCG|51199.3|mIU/mL|1-3|||F||Siemens|20111026133929|20111026132457|VistaDim1|" & vbCrLf
strBuffer = strBuffer & "C|1|I|R001^11103AA|R|" & vbCrLf
strBuffer = strBuffer & "C|2|I|R002^20120412000000|R|" & vbCrLf
strBuffer = strBuffer & "C|3|I|R003^000844|R|" & vbCrLf
strBuffer = strBuffer & "C|4|I|E115^Diluted|E|" & vbCrLf
strBuffer = strBuffer & "C|5|I|E521^Above Reference Range|E|" & vbCrLf
strBuffer = strBuffer & "L|1|N|" & vbCrLf
strBuffer = strBuffer & "A5" & vbCrLf

strBuffer = strBuffer & "" & vbCrLf
    
    
strBuffer = ""
strBuffer = strBuffer & ""
strBuffer = strBuffer & "1H|\^&|||serial HOST|||||||P|DB 100102|20111101140952|" & vbCrLf
strBuffer = strBuffer & "P|1||||^^^|||U|^||||^^^^|||||^|||||\||^|||||^|||||" & vbCrLf
strBuffer = strBuffer & "O|1|1120047241||^^^4015^2202^PBNPM|R||||||N|||20111028105837|S^|^^^^||1|1|||20111101140952|||F||||||" & vbCrLf
strBuffer = strBuffer & "C|1|I|G001^PL000374-1|G|" & vbCrLf
strBuffer = strBuffer & "R|1|^^^4076" & vbCrLf
strBuffer = strBuffer & "215^2202^PBNPM|||5-125|||F||Siemens|20111101140952|20111028122251|VistaDim1|" & vbCrLf
strBuffer = strBuffer & "C|1|I|R001^11138AB|R|" & vbCrLf
strBuffer = strBuffer & "C|2|I|R002^20120517000000|R|" & vbCrLf
strBuffer = strBuffer & "C|3|I|R003^002220|R|" & vbCrLf
strBuffer = strBuffer & "C|4|I|E521^Above Assay Range|E|" & vbCrLf
strBuffer = strBuffer & "L|1|N|" & vbCrLf
strBuffer = strBuffer & "2F" & vbCrLf
strBuffer = strBuffer & "" & vbCrLf
    
    
strBuffer = ""
strBuffer = strBuffer & ""
strBuffer = strBuffer & "1H|\^&|||serial HOST|||||||P|DB 100102|20111107153607|" & vbCrLf
strBuffer = strBuffer & "P|1||||^^^|||U|^||||^^^^|||||^|||||\||^|||||^|||||" & vbCrLf
strBuffer = strBuffer & "O|1|1120030841||^^^1802^2202^RF|R||||||N|||20111107152708|S^|^^^^||1|1|||20111107153607|||F||||||" & vbCrLf
strBuffer = strBuffer & "C|1|I|G001^PL000374-1|G|" & vbCrLf
strBuffer = strBuffer & "R|1|^^^1802^52" & vbCrLf
strBuffer = strBuffer & "22202^RF|||0-15|||F||Siemens|20111107153607|20111107161915|VistaDim1|" & vbCrLf
strBuffer = strBuffer & "C|1|I|R001^11143MA|R|" & vbCrLf
strBuffer = strBuffer & "C|2|I|R002^20120509000000|R|" & vbCrLf
strBuffer = strBuffer & "C|3|I|R003^004156|R|" & vbCrLf
strBuffer = strBuffer & "C|4|I|E521^Above Assay Range|E|" & vbCrLf
strBuffer = strBuffer & "L|1|N|" & vbCrLf
strBuffer = strBuffer & "70" & vbCrLf
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




Private Sub E411(asData As String)
    
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim i As Integer
    Dim ii As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    
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
    X = 0
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
                If vasID.MaxRows < gRow Then
                    vasID.MaxRows = gRow
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
                If vasRes.MaxRows < lsResRow Then
                    vasRes.MaxRows = lsResRow
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
    
'On Error Resume Next
    
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
    
'    GetSetup
    
    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
'        Call ErrWrite("장비와 연결되지 않았습니다")
    End If
    
    If Not Connect_Local Then
        Call ErrWrite("로컬서버와 연결되지 않았습니다")
'        MsgBox "연결되지 않았습니다."
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

    '-- osw 추가
    For i = 1 To 3
        If Not Connect_PRServer Then
            cn_cnt = cn_cnt + 1
            If cn_cnt = 3 Then
                If Not Connect_DRServer Then
                    Call ErrWrite("메인서버와 연결되지 않았습니다")
                    'MsgBox "연결되지 않았습니다."
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
'''    Dim strOutput As String     '송신할 데이터
'''
'''
'''    Select Case intSndPhase
'''        Case 1  '## Header
'''            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
'''            intSndPhase = 2
'''            intFrameNo = intFrameNo + 1
'''        Case 2  '## Patient
'''            strOutput = intFrameNo & "P|1" & vbCr & ETX
'''            intSndPhase = 4
'''            'strOutput = intFrameNo & "P|1|||||||||||||||||||||||||||||||||" & vbCr & ETX
'''            intFrameNo = intFrameNo + 1
'''
'''        Case 3  '## No Order
'''
'''        Case 4  '## Order
'''            If mOrder.NoOrder = True Then
'''                '## 접수정보가 없을경우
'''                strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & _
'''                            "^" & mOrder.TubePos & "^^SAMPLE^NORMAL|ALL" & _
'''                            "|R||||||C||||||||||||||Q" & vbCr & ETX
'''                intSndPhase = 5
'''
'''            Else
'''                If mOrder.IsSending = False Then   '## 최초 보낼때
'''                    strOutput = "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & "^" & mOrder.TubePos & _
'''                                "^^SAMPLE^NORMAL|" & mOrder.Order & "|R||||||N||||||||||||||Q"
'''
'''                                '3O|1|9905300211|1^00014^1^^SAMPLE^NORMAL|ALL|R|20110613090006|||||X||||||||||||||O|||||
'''                                '90
'''                    If Len(strOutput) > 230 Then
'''                        mOrder.IsSending = True
'''                        mOrder.Order = Mid$(strOutput, 231)
'''                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'''                        intSndPhase = 4
'''                    Else
'''                        strOutput = intFrameNo & strOutput & vbCr & ETX
'''                        intSndPhase = 5
'''                    End If
'''                Else                        '## 남은 문자열이 있을때
'''                    strOutput = mOrder.Order
'''                    If Len(strOutput) > 230 Then
'''                        mOrder.Order = Mid$(strOutput, 231)
'''                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'''                        intSndPhase = 4
'''                    Else
'''                        mOrder.IsSending = False
'''                        strOutput = intFrameNo & strOutput & vbCr & ETX
'''                        intSndPhase = 5
'''                    End If
'''                End If
'''            End If
'''            intFrameNo = intFrameNo + 1
'''        Case 5  '## Termianator
'''            strOutput = intFrameNo & "L|1" & vbCr & ETX
'''            intSndPhase = 6
'''            intFrameNo = intFrameNo + 1
'''
'''        Case 6  '## EOT
'''            strState = ""
'''            MSComm1.Output = EOT
'''            Save_Raw_Data "[Tx]" & EOT
'''            intFrameNo = 1
'''
'''            Exit Sub
'''    End Select
'''
'''    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
'''    MSComm1.Output = strOutput
'''    Debug.Print strOutput
'''    Save_Raw_Data "[Tx]" & strOutput

    Dim strOutput As String     '송신할 데이터


    Select Case intSndPhase
        Case 0
            strOutput = EOT
            MSComm1.Output = strOutput
            Save_Raw_Data "[Tx]" & strOutput
            strState = ""
            Exit Sub

        '-- 최초 오더전송
        Case 1  '## Header
            '## Header
            strOutput = "H|\^&||||||||||P|" & vbCr

            '## Patient
            strOutput = strOutput & "P|1|" & vbCr

            '## Order
            If mOrder.NoOrder = False Then
                strOutput = strOutput & "O|1|" & mOrder.BarNo & "||" & mOrder.Order & _
                            "|R||||||||||S||||||||||Q" & vbCr
            Else
                '## 접수정보가 없는경우
                strOutput = strOutput & "O|1|" & mOrder.BarNo & "|||R||||||C||||||||||||||Q" & vbCr
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

    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    MSComm1.Output = strOutput
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
            
'            Buffer = strBuffer
'            lngBufLen = Len(Buffer)
'            Debug.Print Buffer
            
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
                                If intBufCnt = 0 Then
                                    intBufCnt = 1
                                    Erase strRecvData
                                    ReDim Preserve strRecvData(intBufCnt)
                                Else
                                    intBufCnt = intBufCnt + 1
                                    ReDim Preserve strRecvData(intBufCnt)
                                End If
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
                                intPhase = 4
                                MSComm1.Output = ACK
                                Save_Raw_Data "[Tx]" & ACK
'                            Case vbLf
'                                intPhase = 4
'                                MSComm1.Output = ACK
'                                Save_Raw_Data "[Tx]" & ACK
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
    
'    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    If Mid(pBarNo, 1, 2) = "99" Then
        Call Get_Sample_Info_QC(intRow)
    Else
        Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    End If
    
    '-- 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.
    '-- intRow 추가
    strItems = GetEquipExamCode_VISTA(gEquip, pBarNo, intRow)
    
    If strItems <> "" Then
        Call ErrWrite(pBarNo & "의 검사항목이 없습니다")
    End If
    
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
        If vasID.MaxRows < intRow Then
            vasID.MaxRows = intRow
        End If
    End If
    
    Call SetText(vasID, pBarNo, intRow, colBarcode)  '3
    Call SetText(vasID, mResult.RackNo, intRow, colRack)       '4
    Call SetText(vasID, mResult.TubePos, intRow, colPos)         '5
    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    Call Get_Sample_Info(intRow)                        '2,6,7,8,9
    
    gRow = intRow
    

'    gOrderExam = GetOrderExamCode(gEquip, pBarNo)
    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)

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
        If strType = "|" Then
            strType = Mid$(strRcvBuf, 1, 1)
        End If
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
'                strBarNo = Format$(mGetP(strRcvBuf, 3, "|"), String$(15, "#"))
'                strTemp1 = mGetP(strRcvBuf, 4, "|")
'                strSeq = mGetP(strTemp1, 1, "^")
'                strRackNo = Format$(mGetP(strTemp1, 2, "^"), "####")
'                strTubePos = Format$(mGetP(strTemp1, 3, "^"), "##")
'
'                mResult.BarNo = strBarNo
'                mResult.SpcPos = strTubePos & "/" & strRackNo
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
                strBarno = Format$(mGetP(strRcvBuf, 3, "|"), String$(15, "#"))
                
                mResult.BarNo = strBarno
                
                Call SetPatInfo(strBarno)
                               

            Case "R"    '## Result
                '## 장비기준 검사명, 결과, Abnormal Flag
                strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 6, "^")
                strResult = mGetP(strRcvBuf, 4, "|")
                    
                If strResult <> "" Then
                    '## 정성결과 저장
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
                        Call ErrWrite(mResult.BarNo & "의 검사오더가 없습니다")

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
                            strState = ""
                        End If
                    End If
                End If
                
                strState = "R"
                
            Case "C"    '## Comment
                '## Abnormal 결과일때 Comment 저장
                If strState = "R" Then
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strTemp2 = mGetP(strRcvBuf, 5, "|")
                    
                    If strResult = "" And strTemp2 = "E" Then
                        strComm = mGetP(strTemp1, 1, "^")
                        If strIntBase = "PBNPM" And strComm = "E521" Then
                            strComm = ">35000"
                        End If
                        
                        '##  인터페이스 결과 컬렉션의 해당 장비기준 검사명이 존재할때만 Comment를 입력 하도록 수정
                        '========================================================================
                              SQL = "Select examcode, examname, seqno "
                        SQL = SQL & "  From equipexam"
                        SQL = SQL & " Where equipno = '" & gEquip & "' "
                        SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                        SQL = SQL & "   and examcode in (" & gOrderExam & ") "
                        res = db_select_Col(gLocal, SQL)
                        
                        lsResRow = vasRes.DataRowCnt + 1
    
                        If res > 0 Then
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
    
                            
                            SetText vasID, strComm, gRow, colA1c                   '결과
                            SetText vasID, strComm, gRow, colA1c + 1                  'Flag
                            
                            SetText vasRes, strIntBase, lsResRow, colEquipCode       '장비코드
                            SetText vasRes, lsExamCode, lsResRow, colExamCode        '검사코드
                            SetText vasRes, lsExamName, lsResRow, colExamName        '검사명
                            SetText vasRes, strComm, lsResRow, colResult           '결과
                            SetText vasRes, lsSeqNo, lsResRow, colSeq                '순번
                            SetText vasRes, strComm, lsResRow, 7                      'Flag
                            
                            Save_Local_One gRow, lsResRow, "1", strComm
                            
                            strComm = ""
                            lsResult_Buff = ""
                        Else
                                  SQL = "Select examcode, examname, seqno "
                            SQL = SQL & "  From equipexam"
                            SQL = SQL & " Where equipno = '" & gEquip & "' "
                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                            res = db_select_Col(gLocal, SQL)
                            
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
    
                            SetText vasID, strComm, gRow, colA1c + 1                  'Flag
                            SetText vasRes, strComm, lsResRow, 7                      'Flag
                            
                            Save_Local_One gRow, lsResRow, "1", strComm
    
                            strComm = ""
                            lsResult_Buff = ""
                        End If
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
    Dim X As Integer
    
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
    X = 0
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
            If vasID.MaxRows < gRow Then
                vasID.MaxRows = gRow
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
            If vasRes.MaxRows < lsResRow Then
                vasRes.MaxRows = lsResRow
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
    
    res = db_select_Vas(gLocal, SQL, vasRes)
    If res = -1 Then
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

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   Caption         =   " BM6010 Interface "
   ClientHeight    =   11130
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   15225
   BeginProperty Font 
      Name            =   "????ü"
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
   ScaleHeight     =   11130
   ScaleWidth      =   15225
   Begin VB.PictureBox Picture1 
      Align           =   1  '?? ????
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   15165
      TabIndex        =   39
      Top             =   0
      Width           =   15225
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   315
         Left            =   5190
         TabIndex        =   50
         Top             =   30
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "????ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   121307136
         CurrentDate     =   40457
      End
      Begin VB.Label Label1 
         Appearance      =   0  '????
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '????
         Caption         =   "?˻?????"
         BeginProperty Font 
            Name            =   "????"
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
         Left            =   4260
         TabIndex        =   51
         Top             =   90
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '????
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '????
         Caption         =   "BM6010 Interface"
         BeginProperty Font 
            Name            =   "????"
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
         TabIndex        =   43
         Top             =   90
         Width           =   1665
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12180
         Picture         =   "frmInterface.frx":14F5
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13380
         Picture         =   "frmInterface.frx":1A7F
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14670
         Picture         =   "frmInterface.frx":2009
         Top             =   90
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '????
         Caption         =   "Port"
         Height          =   195
         Index           =   0
         Left            =   11640
         TabIndex        =   42
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '????
         Caption         =   "Send"
         Height          =   195
         Left            =   12765
         TabIndex        =   41
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '????
         Caption         =   "Receive"
         Height          =   195
         Left            =   13800
         TabIndex        =   40
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   3375
      Left            =   15180
      TabIndex        =   32
      Top             =   7230
      Visible         =   0   'False
      Width           =   8625
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   240
         TabIndex        =   33
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
         TabIndex        =   34
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
         SpreadDesigner  =   "frmInterface.frx":400C
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   6375
      Left            =   15240
      TabIndex        =   10
      Top             =   660
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame Frame3 
         Height          =   9675
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   14625
         Begin VB.CommandButton cmdRClear 
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11520
            TabIndex        =   76
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdRTrans 
            Caption         =   "????????????"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5250
            TabIndex        =   75
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   74
            Top             =   780
            Width           =   225
         End
         Begin VB.CommandButton cmdRSch 
            Caption         =   "???ð?????ȸ"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3720
            TabIndex        =   72
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "&Excel"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13020
            TabIndex        =   71
            Top             =   240
            Width           =   1395
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   7860
            TabIndex        =   65
            Top             =   630
            Width           =   6675
            Begin VB.Label Label2 
               Caption         =   "???ڵ???ȣ :"
               BeginProperty Font 
                  Name            =   "????ü"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   510
               TabIndex        =   70
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   1
               Left            =   1995
               TabIndex        =   69
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label4 
               Caption         =   "ȯ?ڸ? :"
               BeginProperty Font 
                  Name            =   "????ü"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3540
               TabIndex        =   68
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   1
               Left            =   4590
               TabIndex        =   67
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   66
               Top             =   720
               Width           =   1155
            End
         End
         Begin VB.CheckBox chkSave 
            Caption         =   "????????"
            BeginProperty Font 
               Name            =   "????ü"
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
            TabIndex        =   64
            Top             =   210
            Width           =   795
         End
         Begin VB.OptionButton optSaveResultR 
            Caption         =   "????"
            BeginProperty Font 
               Name            =   "????ü"
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
            TabIndex        =   63
            Top             =   270
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optSaveResultR 
            Caption         =   "????"
            BeginProperty Font 
               Name            =   "????ü"
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
            TabIndex        =   62
            Top             =   270
            Width           =   735
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   8070
            Left            =   7860
            TabIndex        =   61
            Top             =   1455
            Width           =   6675
            _Version        =   393216
            _ExtentX        =   11774
            _ExtentY        =   14235
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????"
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
            SpreadDesigner  =   "frmInterface.frx":4224
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   73
            Top             =   300
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   121307136
            CurrentDate     =   40457
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   8805
            Left            =   165
            TabIndex        =   77
            Top             =   720
            Width           =   7605
            _Version        =   393216
            _ExtentX        =   13414
            _ExtentY        =   15531
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????ü"
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
            SpreadDesigner  =   "frmInterface.frx":7F26
            UserResize      =   2
         End
         Begin VB.Label Label9 
            Appearance      =   0  '????
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '????
            Caption         =   "?˻?????"
            BeginProperty Font 
               Name            =   "????"
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
            TabIndex        =   79
            Top             =   390
            Width           =   780
         End
         Begin VB.Label Label7 
            Appearance      =   0  '????
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '????
            Caption         =   "????????"
            BeginProperty Font 
               Name            =   "????"
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
            TabIndex        =   78
            Top             =   360
            Width           =   780
         End
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   1455
         Left            =   120
         TabIndex        =   24
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
         SpreadDesigner  =   "frmInterface.frx":893D
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   2235
         Left            =   3780
         TabIndex        =   11
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
         SpreadDesigner  =   "frmInterface.frx":8B55
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '????
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         Picture         =   "frmInterface.frx":8D6D
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   25
         Top             =   5790
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   210
         TabIndex        =   23
         Top             =   5640
         Width           =   1215
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "????"
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
         ScrollBars      =   2  '????
         TabIndex        =   17
         Top             =   4830
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   5100
         TabIndex        =   16
         Top             =   5700
         Width           =   645
      End
      Begin VB.TextBox Text_ini 
         BeginProperty Font 
            Name            =   "????"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4440
         TabIndex        =   15
         Top             =   5715
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   3780
         MultiLine       =   -1  'True
         ScrollBars      =   3  '??????
         TabIndex        =   14
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
         Style           =   1  '?׷???
         TabIndex        =   13
         Top             =   5640
         Value           =   1  'Ȯ??
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   4860
         TabIndex        =   12
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
         Begin VB.PictureBox imlStatus 
            BackColor       =   &H80000005&
            Height          =   480
            Left            =   1140
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   82
            Top             =   180
            Width           =   1200
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1485
         Left            =   120
         TabIndex        =   18
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
         SpreadDesigner  =   "frmInterface.frx":92F7
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   2205
         Left            =   3780
         TabIndex        =   19
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
         SpreadDesigner  =   "frmInterface.frx":950F
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1485
         Left            =   120
         TabIndex        =   20
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
         SpreadDesigner  =   "frmInterface.frx":9727
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '????
         BorderStyle     =   1  '???? ????
         BeginProperty Font 
            Name            =   "????"
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
         TabIndex        =   35
         Top             =   5760
         Width           =   825
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2940
         TabIndex        =   22
         Top             =   5730
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3720
         TabIndex        =   21
         Top             =   5730
         Width           =   705
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
      Tabs            =   1
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "????ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "WorkList"
      TabPicture(0)   =   "frmInterface.frx":993F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtTest"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.TextBox txtTest 
         Height          =   375
         Left            =   3900
         TabIndex        =   54
         Top             =   30
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton Command16 
         Caption         =   "?????׽?Ʈ"
         Height          =   435
         Left            =   4590
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   9645
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   14625
         Begin VB.CommandButton Command1 
            Caption         =   "EOT"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10770
            TabIndex        =   81
            Top             =   240
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.CheckBox chkWAll 
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   780
            Width           =   225
         End
         Begin VB.TextBox txtRack 
            Alignment       =   2  '??? ????
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   6360
            TabIndex        =   80
            Text            =   "1"
            Top             =   270
            Width           =   465
         End
         Begin FPSpread.vaSpread vasWorkList 
            Height          =   8805
            Left            =   60
            TabIndex        =   49
            Top             =   720
            Width           =   7785
            _Version        =   393216
            _ExtentX        =   13732
            _ExtentY        =   15531
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            DisplayRowHeaders=   0   'False
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????ü"
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
            SpreadDesigner  =   "frmInterface.frx":995B
            UserResize      =   2
         End
         Begin VB.TextBox txtNum 
            Alignment       =   2  '??? ????
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   6870
            TabIndex        =   58
            Text            =   "0"
            Top             =   270
            Width           =   555
         End
         Begin VB.CommandButton cmdPatDelete 
            Caption         =   "????"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4950
            TabIndex        =   56
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdDownload 
            Caption         =   "Down"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   55
            Top             =   -150
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cboChk 
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frmInterface.frx":A42F
            Left            =   2340
            List            =   "frmInterface.frx":A439
            TabIndex        =   45
            Top             =   -90
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "??ȸ"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3900
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optSaveResult 
            Caption         =   "????"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   9750
            TabIndex        =   37
            Top             =   270
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton optSaveResult 
            Caption         =   "????"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   8970
            TabIndex        =   36
            Top             =   270
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11520
            TabIndex        =   9
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13020
            TabIndex        =   8
            Top             =   240
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Frame Frame6 
            Height          =   585
            Left            =   7860
            TabIndex        =   26
            Top             =   630
            Width           =   6675
            Begin VB.Label Label8 
               Caption         =   "???ڵ???ȣ :"
               BeginProperty Font 
                  Name            =   "????ü"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   510
               TabIndex        =   31
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   0
               Left            =   1995
               TabIndex        =   30
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label6 
               Caption         =   "ȯ?ڸ? :"
               BeginProperty Font 
                  Name            =   "????ü"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3540
               TabIndex        =   29
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   0
               Left            =   4590
               TabIndex        =   28
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   27
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
         Begin FPSpread.vaSpread vasID 
            Height          =   4425
            Left            =   -510
            TabIndex        =   7
            Top             =   5100
            Visible         =   0   'False
            Width           =   7605
            _Version        =   393216
            _ExtentX        =   13414
            _ExtentY        =   7805
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????ü"
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
            SpreadDesigner  =   "frmInterface.frx":A449
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8070
            Left            =   7860
            TabIndex        =   6
            Top             =   1455
            Width           =   6675
            _Version        =   393216
            _ExtentX        =   11774
            _ExtentY        =   14235
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????"
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
            SpreadDesigner  =   "frmInterface.frx":AF4F
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
               Appearance      =   0  '????
               Height          =   1455
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  '????
               TabIndex        =   4
               Top             =   240
               Width           =   5775
            End
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   315
            Left            =   2520
            TabIndex        =   46
            Top             =   300
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   121307137
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   315
            Left            =   1050
            TabIndex        =   47
            Top             =   300
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "????ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   121307137
            CurrentDate     =   40248
         End
         Begin VB.Label Label13 
            Appearance      =   0  '????
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '????
            Caption         =   "Seq"
            BeginProperty Font 
               Name            =   "????"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   7140
            TabIndex        =   59
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label12 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "????"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2370
            TabIndex        =   52
            Top             =   360
            Width           =   105
         End
         Begin VB.Label Label1 
            Appearance      =   0  '????
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '????
            Caption         =   "ó??????"
            BeginProperty Font 
               Name            =   "????"
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
            Left            =   180
            TabIndex        =   48
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label5 
            Appearance      =   0  '????
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '????
            Caption         =   "????????"
            BeginProperty Font 
               Name            =   "????"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   7905
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   780
         End
      End
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  '?Ʒ? ????
      Height          =   405
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   15165
      TabIndex        =   0
      Top             =   10725
      Width           =   15225
   End
   Begin VB.Menu MnMain 
      Caption         =   "Main"
      Begin VB.Menu MnExit 
         Caption         =   "????"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "Setting"
      Begin VB.Menu MnTConfig 
         Caption         =   "???ż???"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "?ڵ弳??"
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

'''vasid, vasrid colum
Const colSpecNo = 0 '?̻???
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
''
''Const colA1c = 12
''Const colIFCC = 13
''Const coleAg = 14
''
''
''
''
'''sendflag
'''0: Order
'''1: Result
'''2: Trans
''
'vasres, vasrres colum
Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
Const colMachResult = 4
Const colResult = 5
Const colSeq = 6
Const colFLAG = 7


'vasid, vasrid colum
'Const colSpecNo = 0 '?̻???
'Const colCheckBox = 1
'Const colBarcode = 2
'Const colRack = 3
'Const colDISK = 3
'Const colPos = 4
'Const colPID = 5
'Const colPName = 6
'Const colSex = 7
'Const colAge = 8
'Const colOCnt = 9
'Const colRCnt = 10
'Const colState = 11
'
'Const colA1c = 12
'Const colIFCC = 13
'Const coleAg = 14




'sendflag
'0: Order
'1: Result
'2: Trans

'vasres, vasrres colum
''Const colEquipCode = 1
''Const colExamCode = 2
''Const colExamName = 3
'''Const colMachResult = 4
''Const colResult = 4
''Const colSeq = 5
''Const colFLAG = 6

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
        MsgBox "?????? ?ڷᰡ ?????ϴ?.", , "?? ??"
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

' Excel Object Library ?? ?????մϴ?.
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
    txtRack = "1"
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
    
    'SELECT ó?? '' ?? üũ?ڽ?
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
            Case "0": SetText vasRID, "????", iRow, colState
            Case "1": SetText vasRID, "????", iRow, colState
            Case "2": SetText vasRID, "?Ϸ?", iRow, colState
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
        
    '-- ??????
    Dim strRequest  As String
    Dim strResponse As String
    Dim varResponse As Variant
    
    Dim strMchNum As String
    Dim strTests  As String
    

    '-- ?˻??????? ????????
                 strRequest = "jobs" + vbTab + "L" + vbTab
    strRequest = strRequest & "hos_org_no" + vbTab + gGINUS_Parm.HCD + vbTab
    strRequest = strRequest & "fr_ymd" + vbTab + pFrDt + vbTab
    strRequest = strRequest & "to_ymd" + vbTab + pToDt + vbTab
    strRequest = strRequest & "mach_cd" + vbTab + "09" + vbTab
    strRequest = strRequest & "smp_no" + vbTab + "%" + vbTab + vbCr
    
    'Debug.Print strRequest

'    strResponse = W2ACALL2("SCC0191A", strRequest, "https://121.78.172.70") '-- ???ڵ??? ?˻????? ??ȸ
    strResponse = W2ACALL2("SCC0191A", strRequest, gGINUS_Parm.URL) '-- ???ڵ??? ?˻????? ??ȸ(https://211.172.17.66)
    
 
'    Debug.Print strResponse
    
    strResponse = Mid(strResponse, 90)
    varResponse = Split(strResponse, vbLf)
    
    For i = 0 To UBound(varResponse) - 1
'        If Trim(Mid(varResponse(0), 1, 10)) <> "0" Then
'            MsgBox "??ȸ?????? ?????ϴ?."
'            Exit Sub
'        End If
        If vasWorkList.DataRowCnt = 0 Then
            vasWorkList.MaxRows = 1
            intRow = 1
        
            txtNum = txtNum + 1
            If txtNum = "85" Then
                txtRack = txtRack + 1
                txtNum = 1
            End If
            
            SetText vasWorkList, "0", intRow, colCheckBox
            'SetText vasWorkList, CStr(txtNum), intRow, colSeqNo
            SetText vasWorkList, Format(txtRack, "00") & "-" & Format(txtNum, "00"), intRow, colSeqNo
            SetText vasWorkList, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colOrdDate
            SetText vasWorkList, mGetP(varResponse(i), 2, vbTab), intRow, colBarcode
            SetText vasWorkList, mGetP(varResponse(i), 6, vbTab), intRow, colPID
            SetText vasWorkList, mGetP(varResponse(i), 7, vbTab), intRow, colPName
            
            Select Case mGetP(varResponse(i), 13, vbTab)
                Case "O": SetText vasWorkList, "?ܷ?", intRow, colRack
                Case "E": SetText vasWorkList, "????", intRow, colRack
                Case "I": SetText vasWorkList, "?Կ?", intRow, colRack
            End Select
            
            strTests = GetOrderExamCode(gEquip, mGetP(varResponse(i), 2, vbTab))
            SetText vasWorkList, strTests, intRow, colPos
            SetText vasWorkList, Len(strTests) / 4, intRow, colOCnt

            chkWAll.Value = 0
        Else
            '-- Same Check
            intRow = getSameRowNum(Trim(mGetP(varResponse(i), 2, vbTab)))
            If intRow = 0 Then
                vasWorkList.MaxRows = vasWorkList.DataRowCnt + 1
                intRow = vasWorkList.MaxRows
            
                txtNum = txtNum + 1
                If txtNum = "85" Then
                    txtRack = txtRack + 1
                    txtNum = 1
                End If
                
                SetText vasWorkList, "0", intRow, colCheckBox
                SetText vasWorkList, Format(txtRack, "00") & "-" & Format(txtNum, "00"), intRow, colSeqNo
                SetText vasWorkList, Mid(mGetP(varResponse(i), 5, vbTab), 1, 8), intRow, colOrdDate
                SetText vasWorkList, mGetP(varResponse(i), 2, vbTab), intRow, colBarcode
                SetText vasWorkList, mGetP(varResponse(i), 6, vbTab), intRow, colPID
                SetText vasWorkList, mGetP(varResponse(i), 7, vbTab), intRow, colPName
                
                Select Case mGetP(varResponse(i), 13, vbTab)
                    Case "O": SetText vasWorkList, "?ܷ?", intRow, colRack
                    Case "E": SetText vasWorkList, "????", intRow, colRack
                    Case "I": SetText vasWorkList, "?Կ?", intRow, colRack
                End Select
                
                strTests = GetOrderExamCode(gEquip, mGetP(varResponse(i), 2, vbTab))
                SetText vasWorkList, strTests, intRow, colPos
                SetText vasWorkList, Len(strTests) / 4, intRow, colOCnt
            
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
'            Case "O": SetText vasWorkList, "?ܷ?", intRow, colRack
'            Case "E": SetText vasWorkList, "????", intRow, colRack
'            Case "I": SetText vasWorkList, "?Կ?", intRow, colRack
'        End Select
    Next
    
    vasWorkList.RowHeight(-1) = 12

End Sub

Private Function getSameRowNum(ByVal strBarno As String) As Integer
Dim i As Integer

    getSameRowNum = 0
    With vasWorkList
        For i = 1 To .DataRowCnt
            .Row = i
            .Col = colBarcode
            If Trim(.Text) = strBarno Then
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

    comEqp.Output = EOT

End Sub

Private Sub imgPort_DblClick()
    
'    '-- ???߽ÿ??? Remark Ǯ? ?׽?Ʈ????
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
        
    strBuffer = ""
    strBuffer = strBuffer & "1H|\^&||||||||||P||05" & vbCrLf
    strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||||||||||3B" & vbCrLf
    strBuffer = strBuffer & "3O|1|11208647111|807^00042^3^^SAMPLE^NORMAL|ALL|R|20111205092128|||||X||||||||||||||O|||||38" & vbCrLf
    strBuffer = strBuffer & "4R|1|^^^321^^0|>100.0|ng/ml|0.000^4.00|>||F|||20111205092406|20111205094226|CF" & vbCrLf
    strBuffer = strBuffer & "5C|1|I|51^Above measuring range|I04" & vbCrLf
    strBuffer = strBuffer & "6R|2|^^^391^^0|13.78|ng/ml|^|N||F|||20111205092448|20111205094308|14" & vbCrLf
    strBuffer = strBuffer & "7L|140" & vbCrLf
    strBuffer = strBuffer & "" & vbCrLf
    
    strBuffer = ""
    strBuffer = strBuffer & "1H|\^&|||iSmart30^iSmart30364^-^1.0.4.1 EX R2||||||||1394-97|2010112514415511" & vbCr
    strBuffer = strBuffer & "2P|1|||||||||||||||||||||||||||||||||2E" & vbCr
    strBuffer = strBuffer & "3O|1||3||||||||||||Sample|||||||||||||||80" & vbCr
    strBuffer = strBuffer & "4R|1|^^^Na+^M|121|mmol/L||N||||||20101125143057|77" & vbCr
    strBuffer = strBuffer & "5R|2|^^^K+^M|1.9|mmol/L||N|||||||59" & vbCr
    strBuffer = strBuffer & "6R|3|^^^Cl-^M|73|mmol/L||N|||||||93" & vbCr
    strBuffer = strBuffer & "7R|4|^^^Hct^M|Out of Range(L)|%||L|||||||38" & vbCr
    strBuffer = strBuffer & "0L|1|NF693" & vbCr
    strBuffer = strBuffer & "" & vbCrLf
        
    strBuffer = ":s  22022                3 0228141104 8 6    41  7    33  8    33 10   243 11    93 13   0.7 15   294 16    44 F2"
    strBuffer = ":S  26026                3 0304141401 3 1   139 10  38.2 11   1.3 D3"
    
    strBuffer = ":s  38038                3 031314073610 1   6.3  2   2.4  3   0.8  6   126  7    54  8    23  9   299 13   1.1 14  19.7F25  4.69 49"
    
    strBuffer = ":S  35035                3 0313140736 9 1   5.2  2   3.6  3   0.3  7    28  8    24  9   127 13   0.7 14   9.7 25  0.93 67"
    
    
    strBuffer = ""
    strBuffer = strBuffer & "1Q 010101101-01         " & vbCr
    strBuffer = strBuffer & ""
    
    Call comEqp_OnComm
        

End Sub



Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
On Error Resume Next
    
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
    
    comEqp.CommPort = gSetup.gPort
    comEqp.RTSEnable = gSetup.gRTSEnable
    comEqp.DTREnable = gSetup.gDTREnable
    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If
    
    If comEqp.PortOpen Then
        frmInterface.StatusBar1.Panels(2).Text = "???? ?Ǿ????ϴ?"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        lblStatus = "?۾???.."
    Else
        frmInterface.StatusBar1.Panels(2).Text = "???? ???? ?ʾҽ??ϴ?"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
'        lblStatus = "?۾? ??????.."
    End If

    If Not Connect_Local Then
        MsgBox "???????? ?ʾҽ??ϴ?."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
''    '-- osw ?߰?
''    For i = 1 To 1
''        If Not Connect_PRServer Then
''            'Cn_Cnt = Cn_Cnt + 1
''            'If Cn_Cnt = 3 Then
''            '    If Not Connect_DRServer Then
''                    MsgBox "???????? ?ʾҽ??ϴ?."
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
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -1), "yyyymmdd")
    
    SQL = "delete from PATRESULT where examdate < '" & sDate & "'"
    Res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
    If lblUser.Caption = "" Then
        Call picLogin_Click
    End If
    'MsgBox "61"
    
    stInterface.Tab = 0

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
'    vasWorkList.MaxRows = 10
    
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

Private Sub Form_Unload(Cancel As Integer)
    If comEqp.PortOpen = True Then
        comEqp.PortOpen = False
    End If

'    Call dce_close_env      ' Server?? ?????? ???? ??
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
'   ???? : ???????? ????
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim intRow  As Integer
    Dim strOutput As String     '?۽??? ??????

On Error Resume Next

    With vasWorkList
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = colCheckBox
            If .Value = "1" Then
                
                '-- ???ڵ?
                '1O 0101030N01014016942                                              U   20140529 1.0111  M10 M11 M12 M13 M14 M15 M16 M17 M18 M19 M2  M20 M21 M22 M23 M24 M25 M26 M27 M28 M29 M3  M30 M4  M5  M6  M7  M8  M9  M CD
                'strOutput = intFrameNo & "O 0101" & Right("000" & GetText(vasID, intRow, colPos), 3) & "N0" & Left(GetText(vasID, intRow, colBarcode) & Space$(13), 13) & _
                            Space$(11) & Space$(32) & "U" & Space(3) & Format(Now, "yyyymmdd") & " 1.0" & "1" & "1" & GetText(vasID, intRow, colRack) & Space(1) & ETX
                
                
                '-- Pos
                '2O 0101018N0150210-01-28 01-28      00077694        ??????          U   20150210 1.01122 M11 M1  M2  M3  M7  M4  M5  M16 M17 M12 M30 M31 M32 M6  M13 M14 M15 M 07
                '1O 0101000N015-0107453-1 01-01      10014350                        U   20151216 1.01111 M13 M14 M4  M5  M6  M8  M9  M E0

                            strOutput = intFrameNo & "O 0101" & Right("000" & GetText(vasWorkList, intRow, colOCnt), 3) & "N0"
                strOutput = strOutput & Left(GetText(vasWorkList, intRow, colBarcode) & Space$(13), 13)
                strOutput = strOutput & Left(GetText(vasWorkList, intRow, colSeqNo) & Space(11), 11)
                strOutput = strOutput & Left(GetText(vasWorkList, intRow, colPID) & Space(16), 16)
                'strOutput = strOutput & LeftB(GetText(vasWorkList, intRow, colPName) & Space(16), 16)
                strOutput = strOutput & Space(16)
                strOutput = strOutput & "U" & Space(3) & Format(Now, "yyyymmdd") & " 1.0" & "1" & "1" & GetText(vasWorkList, intRow, colPos) & Space(1) & ETX
                intFrameNo = intFrameNo + 1
                        
                
                strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
                'strOutput = STX & strOutput & vbCrLf
                comEqp.Output = strOutput
                Debug.Print strOutput
                SetRawData "[Tx]" & strOutput
                .Row = intRow
                .Col = colState
                .Value = "Order"
                
                .Row = intRow
                .Col = colCheckBox
                .Value = "0"
                                
                SetBackColor vasWorkList, intRow, intRow, 1, colState, 202, 255, 212
                'Sleep (100)
                
                'strState = "E"
                'intFrameNo = 1
                
                Exit For
            End If
        Next
            
        If intFrameNo >= 8 Then
            intFrameNo = 0
        End If
        
        If intRow = vasWorkList.MaxRows Then
            strState = "E"
            intFrameNo = 1
            Exit Sub
        End If
    
        For intRow = 1 To vasWorkList.MaxRows
            .Row = intRow
            .Col = colCheckBox
            If .Value = "1" Then
                Exit Sub
            End If
        Next
        
        strState = "E"
        intFrameNo = 1
    
    End With
    
End Sub

''''-----------------------------------------------------------------------------'
''''   ???? : ???????? ????
''''-----------------------------------------------------------------------------'
'''Private Sub SendOrder()
'''    Dim strOutput As String     '?۽??? ??????
'''
'''    '-- ASTM TYPE?? Define ?ؾ???.
'''    '-- ASTM TYPE = Standard
'''    If gASTMFormat = "1" Then
'''        Select Case intSndPhase
'''            Case 1  '## Header
'''                strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
'''                intSndPhase = 2
'''                intFrameNo = intFrameNo + 1
'''            Case 2  '## Patient
'''                strOutput = intFrameNo & "P|1" & vbCr & ETX
'''                intSndPhase = 4
'''                intFrameNo = intFrameNo + 1
'''
'''            Case 3  '## No Order
'''
'''            Case 4  '## Order
'''                If mOrder.NoOrder = True Then
'''
'''                    '## ?????????? ????????
'''                    strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & "^" & mOrder.TubePos & "^^SAMPLE^NORMAL|ALL" & _
'''                                "|R||||||C||||||||||||||Q" & vbCr & ETX
'''                    intSndPhase = 5
'''
'''                Else
'''                    If mOrder.IsSending = False Then   '## ???? ??????
'''                        strOutput = "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & "^" & mOrder.TubePos & "^^SAMPLE^NORMAL|" & mOrder.Order & "|R||||||N||||||||||||||Q"
'''
'''                        If Len(strOutput) > 230 Then
'''                            mOrder.IsSending = True
'''                            mOrder.Order = Mid$(strOutput, 231)
'''                            strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'''                            intSndPhase = 4
'''                        Else
'''                            strOutput = intFrameNo & strOutput & vbCr & ETX
'''                            intSndPhase = 5
'''                        End If
'''                    Else                        '## ???? ???ڿ??? ??????
'''                        strOutput = mOrder.Order
'''                        If Len(strOutput) > 230 Then
'''                            mOrder.Order = Mid$(strOutput, 231)
'''                            strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'''                            intSndPhase = 4
'''                        Else
'''                            mOrder.IsSending = False
'''                            strOutput = intFrameNo & strOutput & vbCr & ETX
'''                            intSndPhase = 5
'''                        End If
'''                    End If
'''                End If
'''                intFrameNo = intFrameNo + 1
'''            Case 5  '## Termianator
'''                strOutput = intFrameNo & "L|1" & vbCr & ETX
'''                intSndPhase = 6
'''                intFrameNo = intFrameNo + 1
'''
'''            Case 6  '## EOT
'''                strState = ""
'''                comEqp.Output = EOT
'''                SetRawData "[Tx]" & EOT
'''                intFrameNo = 1
'''
'''                Exit Sub
'''        End Select
'''    '-- ASTM TYPE = Long [=VISTA 500, Hitachi, Modular]
'''    ElseIf gCOMFormat = "2" Then
'''        Select Case intSndPhase
'''            Case 0
'''                strOutput = EOT
'''                comEqp.Output = strOutput
'''                'Save_Raw_Data "[Tx]" & strOutput
'''                strState = ""
'''                Exit Sub
'''
'''            Case 1  '## Header
'''                '## Header
'''                strOutput = "H|\^&||||||||||P|" & vbCr
'''
'''                '## Patient
'''                strOutput = strOutput & "P|1|" & vbCr
'''
'''                '## Order
'''                If mOrder.NoOrder = False Then
'''                    strOutput = strOutput & "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||S||||||||||Q" & vbCr
'''
'''                    Select Case gOPTVersion
'''                    Case "1.0"  '## Version 1.0
'''                                'strOutput = strOutput & "O|1|0^" & Format$(mOrder.BarNo, String$(13, "@")) & "^" & mOrder.SpcType & "^" & mOrder.RackNo & "^" & mOrder.Pos & "|" & _
'''                                                        mOrder.Kind & "|" & mOrder.GetOrder & "|" & mOrder.Priority & "||||||N||^^||||||^^^^||||||O" & vbCr
'''                    Case "1.3"  '## Version 1.3
'''                                'strOutput = strOutput & "O|1|" & mOrder.BarNo & "|" & mOrder.GetInstSpcId & "|" & mOrder.GetOrder & "|" & mOrder.GetPriority & _
'''                                                        "||||||N||||" & mOrder.GetSampleType & "||||||||||O" & vbCr
'''                    End Select
'''
'''                Else
'''                    '## ?????????? ???°???: ?˻??׸? ?????? ?????? ????!
'''                    strOutput = strOutput & "O|1|" & mOrder.BarNo & "|||R||||||C||||||||||||||Q" & vbCr
'''
'''                    Select Case gOPTVersion
'''                    Case "1.0"  '## Version 1.0
'''                                'strOutput = strOutput & "O|1|0^" & Format$(mOrder.BarNo, String$(13, "@")) & "^" & mOrder.SpcType & "^" & mOrder.RackNo & "^" & mOrder.Pos & "|" & _
'''                                                        mOrder.Kind & "||" & mOrder.Priority & "||||||N||^^||||||^^^^||||||O" & vbCr
'''                    Case "1.3"  '## Version 1.3
'''                                'strOutput = strOutput & "O|1|" & mOrder.BarNo & "|" & mOrder.GetInstSpcId & "||R" & _
'''                                                        "||||||N||||" & mOrder.GetSampleType & "||||||||||O" & vbCr
'''                    End Select
'''
'''                End If
'''
'''                '## Termianator
'''                strOutput = strOutput & "L|1|N" & vbCr
'''                strOutput = intFrameNo & strOutput
'''            Case 2
'''
'''        End Select
'''
'''        If Len(strOutput) >= 230 Then
'''            mOrder.Order = Mid$(strOutput, 231)
'''            strOutput = Mid$(strOutput, 1, 230) & ETB
'''            intSndPhase = 2
'''        Else
'''            strOutput = strOutput & ETX
'''            intSndPhase = 0
'''        End If
'''
'''    End If
'''
'''    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
'''    comEqp.Output = strOutput
'''    Debug.Print strOutput
'''    SetRawData "[Tx]" & strOutput
'''
'''End Sub

'-----------------------------------------------------------------------------'
'   ???? : ?ش? ???ڿ??? CheckSum?? ????
'   ?μ? :
'       - pMsg : ???ڿ?
'   ??ȯ : CheckSum
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

'-- ???ݳ?¥?? ?˻????? ?????Ѵ?
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
           ' Buffer = strBuffer
            
'            txtData = txtData & Buffer
            
            SetRawData "[Rx]" & Buffer
            lngBufLen = Len(Buffer)
            
            Debug.Print Buffer

            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case BufChar
                    Case STX
                        intBufCnt = intBufCnt + 1
                        ReDim Preserve strRecvData(intBufCnt)
                    Case ETB
                    Case ENQ
                        comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                        strState = ""
                        intFrameNo = 1
                    Case ACK
                        If strState = "E" Then
                            comEqp.Output = EOT
                            SetRawData "[Tx]" & EOT
                            strState = ""
                        End If
                        
                        If strState = "Q" Then
                            Call SendOrder
                        End If

                    Case NAK
                        If strState = "Q" Then
                            Call SendOrder
                        End If
                    Case ETX
'                        Call EditRcvData
                    Case EOT
                        Call EditRcvData
                        Erase strRecvData
                        
'                        If strState = "Q" Then
'                            comEqp.Output = ENQ
'                            SetRawData "[Tx]" & ENQ
'                        End If
                    Case vbCr
                    Case vbLf
                        comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case Else
                        If intBufCnt >= 1 Then
                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                        End If
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
            EVMsg$ = "CTS ???? ????"
        Case comEvDSR
            EVMsg$ = "DSR ???? ????"
        Case comEvCD
            EVMsg$ = "CD ???? ????"
        Case comEvRing
            EVMsg$ = "??ȭ ???? ?︮?? ??"
        Case comEvEOF
            EVMsg$ = "EOF ????"

        '???? ?޽???
        Case comBreak
            ERMsg$ = "?ߴ? ??ȣ ????"
        Case comCDTO
            ERMsg$ = "?ݼ??? ???? ?ð? ?ʰ?"
        Case comCTSTO
            ERMsg$ = "CTS ?ð? ?ʰ?"
        Case comDCB
            ERMsg$ = "DCB ?˻? ????"
        Case comDSRTO
            ERMsg$ = "DSR ?ð? ?ʰ?"
        Case comFrame
            ERMsg$ = "?????̹? ????"
        Case comOverrun
            ERMsg$ = "?и?Ƽ ????"
        Case comRxOver
            ERMsg$ = "???? ???? ?ʰ?"
        Case comRxParity
            ERMsg$ = "?и?Ƽ ????"
        Case comTxFull
            ERMsg$ = "???? ???ۿ? ?????? ????"
        Case Else
            ERMsg$ = "?? ?? ???? ???? ?Ǵ? ?̺?Ʈ"
    End Select


End Sub

'-----------------------------------------------------------------------------'
'   ???? : ?ش? ???ڵ???ȣ?? ???? ???????? ??ȸ, ǥ??, ?˻???????????
'   ?μ? :
'       - pBarNo : ???ڵ???ȣ
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    
    intRow = -1
    
    For i = 1 To vasWorkList.DataRowCnt
        If Trim(GetText(vasWorkList, i, colBarcode)) = pBarNo Then
            intRow = i
            gRow = intRow
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasWorkList.DataRowCnt + 1
        If vasWorkList.MaxRows < intRow Then
            vasWorkList.MaxRows = intRow
            gRow = intRow
        End If
    End If
    
    '-- ???????????? ǥ??
'    Call SetText(vasID, pBarNo, intRow, colBarcode)         '3  ???ڵ?
'    Call SetText(vasID, mOrder.RackNo, intRow, colRack)     '4  Rack??ȣ
'    Call SetText(vasID, mOrder.TubePos, intRow, colPos)     '5  Pos??ȣ
'    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    
    '-- ?˻??? ???? ???????̺????? ?????? ǥ??(for ??ũ????Ʈ)  '6,7,8,9
    'Call GetSampleInfoW(intRow)
    
    '-- ???ڵ???ȣ?? ?????ϴ? ?˻??ڵ? ????????(?μ? : ?????ڵ?,???ڵ???ȣ)
    'gOrderExam = GetOrderExamCode(gEquip, pBarNo)
    
    '-- ???????̺????? ?˻??׸??? ?ش??ϴ? ?˻?ä?? ã?ƿ??? (intRow = ???? ?˻??ߴ? ???ڵ尡 ?ٽ? ?ö??? ???? ??ġ?? ??ã?´?.)
    'strItems = GetGetEquipExamCode_CentaurCP(gEquip, pBarNo, intRow)

    '-- ?˻?ä?η? ???????? ??????
'    If Trim(strItems) = "" Then
'        mOrder.NoOrder = True
'        mOrder.Order = strItems
'    Else
'        mOrder.NoOrder = False
'        mOrder.Order = strItems
'    End If
    
    strItems = GetText(vasWorkList, intRow, colPos)
    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = ""
        'HOST->1O 0101009N0100000029738                                            U   20130704 1.0114  M6  M2  M3  M11 M5  M1  M8  M7  M AC
        Call SetText(vasWorkList, "0", intRow, colCheckBox)         '1
        'Call SetText(vasWorkList, strItems, intRow, colRack)
        'Call SetText(vasWorkList, Len(strItems) / 4, intRow, colPos)
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
        'HOST->1O 0101009N0100000029738                                            U   20130704 1.0114  M6  M2  M3  M11 M5  M1  M8  M7  M AC
        Call SetText(vasWorkList, "1", intRow, colCheckBox)         '1
        'Call SetText(vasWorkList, strItems, intRow, colRack)
        'Call SetText(vasWorkList, Len(strItems) / 4, intRow, colPos)
    End If
    
    
    Call SetText(vasWorkList, "Order", intRow, colState)         '12 ????????

End Sub

'-----------------------------------------------------------------------------'
'   ???? :
'   ?μ? :
'       - pBarNo : ???ڵ???ȣ
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    
    intRow = -1
    For i = 1 To vasWorkList.DataRowCnt
        If Trim(GetText(vasWorkList, i, colSeqNo)) = pBarNo Then
            intRow = i
            
            '-- ???ڵ???ȣ?? ?????ϴ? ?˻??ڵ? ????????(?μ? : ?????ڵ?,???ڵ???ȣ)
            gOrderExam = GetOrderExamCode(gEquip, pBarNo)
            
            SetText vasWorkList, mGetP(gOrderExam, 2, "@"), intRow, colPos
            
            gOrderExam = mGetP(gOrderExam, 1, "@")
            Exit For
        End If
    Next i
    
'    If intRow < 0 Then
'        intRow = vasWorkList.DataRowCnt + 1
'        If vasWorkList.MaxRows < intRow Then
'            vasWorkList.MaxRows = intRow
'        End If
'    End If
'
'    '-- ???????????? ǥ??
'    Call SetText(vasWorkList, pBarNo, intRow, colBarcode)             '2 Barcode
'    Call SetText(vasWorkList, mResult.RackNo, intRow, colRack)        '3 Rack
'    Call SetText(vasWorkList, mResult.TubePos, intRow, colPos)        '4 Pos
'    Call vasActiveCell(vasWorkList, intRow, colBarcode)
'
'    '-- ???????????? ??????
'    Call ClearSpread(vasRes)
'
'    '-- ?˻??? ???? ???????̺? ?????? ǥ??(for ??ũ????Ʈ)  '5,6,7,8
'    Call GetSampleInfoW(intRow)                                '5,6,7,8
    
    '-- ???? Row
    gRow = intRow
    

End Sub


''''-----------------------------------------------------------------------------'
''''   ???? : ?????κ? ?????? ?????? ????
''''-----------------------------------------------------------------------------'
'''Private Sub EditRcvDataASTM()
'''    Dim strRcvBuf    As String   '?????? Data
'''    Dim strType      As String   '?????? Record Type
'''    Dim strBarno     As String   '?????? ???ڵ???ȣ
'''    Dim strSeq       As String   '?????? Sequence
'''    Dim strRackNo    As String   '?????? Rack Or Disk No
'''    Dim strTubePos   As String   '?????? Tube Position
'''    Dim strIntBase   As String   '?????? ???????? ?˻???
'''    Dim strResult    As String   '?????? ????(????)
'''    Dim strIntResult As String   '?????? ????(????)
'''    Dim strQCResult  As String   '?????? ????(QC)
'''    Dim strFlag      As String   '?????? Abnormal Flag
'''    Dim strComm      As String   '?????? Comment
'''    Dim strTemp1     As String
'''    Dim strTemp2     As String
'''    Dim intCnt       As Integer
'''
'''    Dim lsExamCode As String
'''    Dim lsExamName As String
'''    Dim lsSeqNo As String
'''    Dim lsResult_Buff As String
'''    Dim lsExamDate As String
'''    Dim lsEquipRes As String
'''    Dim lsResRow    As String
'''    Dim ii As Integer
'''    Dim strTmp      As String
'''    Dim intIdx      As Integer
'''    Dim varRcvBuf   As Variant
'''    Dim intRow      As Integer
'''    Dim i As Integer
'''
'''    'strRcvBuf = strRecvData(1)
'''    'varRcvBuf = Split(strRcvBuf, vbCr)
'''
'''On Error Resume Next
'''
'''    For intCnt = 1 To UBound(strRecvData)
'''        strRcvBuf = strRecvData(intCnt)
'''        strBarno = ""
'''        strBarno = Trim(Mid(strRcvBuf, 4, 3))
'''
'''        For i = 1 To vasWorkList.DataRowCnt
'''            vasWorkList.Row = i
'''            vasWorkList.Col = colSeqNo
'''            If Trim(vasWorkList.Text) = strBarno Then
'''                vasWorkList.Col = colBarcode
'''                strBarno = vasWorkList.Text
'''                gRow = i
'''                Exit For
'''            End If
'''        Next
'''
'''        If strBarno = "" Then Exit Sub
'''
''''        With mResult
''''            .BarNo = strBarNo
''''        End With
'''
'''
'''        If gRow < 0 Then
'''            Exit Sub
'''        End If
'''
'''
'''        For i = 40 To Len(strRcvBuf) Step 9
''''            Case "R"    '## Result
'''            '## ???????? ?˻???, ????, Abnormal Flag
'''            strIntBase = Trim(Mid(strRcvBuf, i, 2))
'''            strResult = Trim(Mid(strRcvBuf, i + 2, 7))
'''
''''            Debug.Print strIntBase & "," & strResult
'''            'strResult = mGetP(strResult, 1, "^")
'''
'''            If strResult <> "" Then
'''                SQL = ""
'''                SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'''                SQL = SQL & "  FROM EQPMASTER"
'''                SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'''                SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'''                'SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
'''                'SQL = SQL & "   AND EXAMCODE in ('C3791','C3792','C3793') "
'''                Res = GetDBSelectColumn(gLocal, SQL)
'''
'''                '-- ???? ???? ????
'''                If Res > 0 Then
'''                    lsExamCode = Trim(gReadBuf(0))
'''                    lsExamName = Trim(gReadBuf(1))
'''                    lsSeqNo = Trim(gReadBuf(2))
'''
'''                    lsResRow = vasRes.DataRowCnt + 1
'''                    If vasRes.MaxRows < lsResRow Then
'''                        vasRes.MaxRows = lsResRow
'''                    End If
'''
'''                    '?Ҽ??? ó??, ???? ???? ó??
'''                    lsEquipRes = strResult
'''                    'strResult = SetResult(strResult, strIntBase)
'''                    lsResult_Buff = strResult
'''
'''                    '-- Work List
'''                    SetText vasID, "Result", gRow, colState                 '11 ????????
'''
'''
'''                    '-- ???? List
'''                    SetText vasRes, strIntBase, lsResRow, colEquipCode      '?????ڵ?
'''                    SetText vasRes, lsExamCode, lsResRow, colExamCode       '?˻??ڵ?
'''                    SetText vasRes, lsExamName, lsResRow, colExamName       '?˻???
'''                    SetText vasRes, lsEquipRes, lsResRow, colMachResult     '????????
'''                    SetText vasRes, strResult, lsResRow, colResult          '????
'''                    SetText vasRes, lsSeqNo, lsResRow, colSeq               '????
'''                    SetText vasRes, strComm, lsResRow, 7                    'Flag
'''                    '-- ???? ????
'''                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                    lsResult_Buff = ""
'''
'''                    strState = "R"
'''
'''                '-- ???? ???? ????
'''                Else
'''
'''                          SQL = "Select examcode, examname, seqno "
'''                    SQL = SQL & "  From EQPMASTER"
'''                    SQL = SQL & " Where equipno = '" & gEquip & "' "
'''                    SQL = SQL & "   and equipcode = '" & strIntBase & "' "
'''                    Res = GetDBSelectColumn(gLocal, SQL)
'''
'''                    If Res > 0 Then
'''                        lsExamCode = Trim(gReadBuf(0))
'''                        lsExamName = Trim(gReadBuf(1))
'''                        lsSeqNo = Trim(gReadBuf(2))
'''
'''                        lsResRow = vasRes.DataRowCnt + 1
'''                        If vasRes.MaxRows < lsResRow Then
'''                            vasRes.MaxRows = lsResRow
'''                        End If
'''
'''                        '?Ҽ??? ó??, ???? ???? ó??
'''                        lsEquipRes = strResult
'''                        'strResult = SetResult(strResult, strIntBase)
'''                        lsResult_Buff = strResult
'''
'''                        '-- Work List
'''                        SetText vasID, "Result", gRow, colState                 '????????
'''
'''                        '-- ???? List
'''                        SetText vasRes, strIntBase, lsResRow, colEquipCode      '?????ڵ?
'''                        SetText vasRes, lsExamCode, lsResRow, colExamCode       '?˻??ڵ?
'''                        SetText vasRes, lsExamName, lsResRow, colExamName       '?˻???
'''                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     '????????
'''                        SetText vasRes, strResult, lsResRow, colResult          '????
'''                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '????
'''                        SetText vasRes, strComm, lsResRow, colFLAG              'Flag
'''                        '-- ???? ????
'''                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                        lsResult_Buff = ""
'''
'''                    End If
'''                End If
'''            End If
''''            Case "C"    '## Comment
'''            '## Abnormal ?????϶? Comment ????
''''                If strFlag <> "N" Then
''''                    strTemp1 = mGetP(strRcvBuf, 4, "|")
''''                    strComm = "[Flag]: " & mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
''''                End If
'''
''''            Case "L"    '## Terminator
'''        Next
'''
'''        '## DB?? ????????
'''        If MnTransAuto.Checked = True And strState = "R" Then
'''
'''            Res = SaveTransDataW(gRow)
'''
'''            If Res = -1 Then
'''                '-- ???? ????
'''                SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
'''                SetText vasID, "Failed", gRow, colState
'''            Else
'''                '-- ???? ????
'''                SetBackColor vasWorkList, gRow, gRow, 1, colState, 202, 255, 112
'''                SetText vasWorkList, "Trans", gRow, colState
'''                SetText vasWorkList, "0", gRow, colCheckBox
'''
'''                SQL = " Update PATRESULT Set " & vbCrLf & _
'''                      " sendflag = '2' " & vbCrLf & _
'''                      " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''                      " And barcode = '" & Trim(GetText(vasWorkList, gRow, colBarcode)) & "' "
'''
'''                Res = SendQuery(gLocal, SQL)
'''                If Res = -1 Then
'''                    SaveQuery SQL
'''                    Exit Sub
'''                End If
'''            End If
'''        End If
'''
'''        'SetText vasID, "Result", gRow, colState
'''        strState = ""
'''
''''        End Select
'''    Next
'''
'''End Sub


'-----------------------------------------------------------------------------'
'   ???? : ?????κ? ?????? ?????? ????
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim strRcvBuf    As String   '?????? Data
    Dim strType      As String   '?????? Record Type
    Dim strBarno     As String   '?????? ???ڵ???ȣ
    Dim strSeq       As String   '?????? Sequence
    Dim strRackNo    As String   '?????? Rack Or Disk No
    Dim strTubePos   As String   '?????? Tube Position
    Dim strIntBase   As String   '?????? ???????? ?˻???
    Dim strResult    As String   '?????? ????
    Dim strQCResult  As String   '?????? ????(QC)
    Dim strFlag      As String   '?????? Abnormal Flag
    Dim strComm      As String   '?????? Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim intCnt       As Integer
    
    Dim lsExamCode   As String
    Dim lsExamName   As String
    Dim lsSeqNo      As String
    Dim lsResult_Buff As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    Dim lsResRow    As String
    Dim ii As Integer
    Dim strTmp      As String
    Dim intIdx      As Integer
    
    Dim intQ        As Integer
    Dim varRcvBuf   As Variant
    Dim strNB       As String
    
    Dim strFe       As String
    Dim strUIBC      As String
    Dim i As Integer

On Error Resume Next

    strFe = ""
    strUIBC = ""
    
    
    For intCnt = 0 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        strType = Mid$(strRcvBuf, 2, 1)
        
        Select Case strType
            Case "Q"    '## Inquiry Order
                strState = ""
                strRcvBuf = Mid(strRcvBuf, 11)
                'varRcvBuf = Split(strRcvBuf, Space$(3))
                '1Q 010101101-01         08
'    strBuffer = strBuffer & "1Q 010101101-01         " & vbCr
'    strBuffer = strBuffer & ""
                
                For intQ = 1 To Len(strRcvBuf) Step 13
                    varRcvBuf = Trim(Mid(strRcvBuf, 1, 13))
                    If Trim(varRcvBuf) = "" Then
                        Exit For
                    Else
                        strBarno = varRcvBuf
                        'strBarno = mGetP(strBarno, 2, "-")
                        For i = 1 To vasWorkList.DataRowCnt
                            vasWorkList.Row = i
                            vasWorkList.Col = colSeqNo
                            If Trim(Trim(vasWorkList.Text)) = Trim(strBarno) Then
                                vasWorkList.Col = colBarcode
                                strBarno = vasWorkList.Text
                                gRow = i
                                Exit For
                            End If
                        Next
                
                        If strBarno <> "" Then
                            Call GetOrder(strBarno)
                            strRcvBuf = Mid(strRcvBuf, 14)
                            strState = "Q"
                        End If
                    End If
                Next
            
            '1R 010101020130704122319N0100000029745                                                            U  020130704 1.011  1M      33          2M      34          3M     156          4M     150          5M     1.1          6M     7.1          7M     4.6          8M     102         10M     0.9         11M     191        28
            
            '1R 010101720151222163459N115-0110028-1 01-01                                                      M   20151222 1.011001M     7.2        002M     4.3        003
            
            Case "R"    '## Result
                strBarno = Trim$(Mid$(strRcvBuf, 27, 13))
                mResult.BarNo = strBarno
                
                If strBarno = "" Then Exit Sub
                                
                strNB = Val(Mid(strRcvBuf, 6, 2))

                If strNB = 1 Then
                    strTmp = Mid$(strRcvBuf, 117)
                    'strTmp = Mid$(strRcvBuf, 114)
                Else
                    strTmp = Mid$(strRcvBuf, 51)
                End If
                
                'Call SetPatInfo(strBarno)
                Call GetOrder(strBarno)
                
                Call GetOrderExamCode(gEquip, strBarno)
                
                Do While Len(strTmp) >= 20
                    
                    strIntBase = Trim(Mid$(strTmp, 1, 3))
                    strResult = Trim(Mid$(strTmp, 5, 15))
                    'strComm = Mid$(strTmp, 10, 1)
                    strResult = Replace(strResult, "h", "")
                    strResult = Replace(strResult, "l", "")
                    
                    '-- Fe
                    If strIntBase = "23" Then
                        strFe = strResult
                    End If
                    
                    '-- UIBC
                    If strIntBase = "24" Then
                        strUIBC = strResult
                    End If

Rst:
                    
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM eqpmaster"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        '-- ???? ???? ????
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '?Ҽ??? ó??, ???? ???? ó??
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            'SetText vasID, strResult, gRow, colA1c                  '????
                            'SetText vasID, strComm, gRow, colA1c + 1                'Flag
                            'SetText vasID, "Result", gRow, colState                 '????????
                            '-- ???? List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '?????ڵ?
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '?˻??ڵ?
                            SetText vasRes, lsExamName, lsResRow, colExamName       '?˻???
                            SetText vasRes, strResult, lsResRow, colResult          '????
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '????
                            SetText vasRes, strComm, lsResRow, 7                    'Flag
                            '-- ???? ????
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            
                            strState = "R"
                        
                        '-- ???? ???? ????
                        Else
                        
                                  SQL = "Select examcode, examname, seqno "
                            SQL = SQL & "  From eqpmaster"
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
                                
                                '?Ҽ??? ó??, ???? ???? ó??
                                lsEquipRes = strResult
                                strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                '-- Work List
                                'SetText vasID, strResult, gRow, colA1c                  '????
                                'SetText vasID, strComm, gRow, colA1c + 1                'Flag
                                SetText vasID, "Result", gRow, colState                 '????????
                                '-- ???? List
                                SetText vasRes, strIntBase, lsResRow, colEquipCode      '?????ڵ?
                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '?˻??ڵ?
                                SetText vasRes, lsExamName, lsResRow, colExamName       '?˻???
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult          '????
                                SetText vasRes, strResult, lsResRow, colResult          '????
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '????
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                '-- ???? ????
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                strState = ""
                            End If
                        End If
                    End If
                    strTmp = Mid$(strTmp, 21)
                Loop
                
                strState = "R"
            
                If MnTransAuto.Checked = True Then
                    
                    Res = SaveTransDataW(gRow)
                    
                    If Res = -1 Then
                        '-- ???? ????
                        SetForeColor vasWorkList, gRow, gRow, 1, colState, 255, 0, 0
                        SetText vasWorkList, "Failed", gRow, colState
                    Else
                        '-- ???? ????
                        SetBackColor vasWorkList, gRow, gRow, 1, colState, 202, 255, 112
                        SetText vasWorkList, "Trans", gRow, colState
                        
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
                    
                    Call vasWorkList_Click(2, gRow)
                End If
            
                SetText vasWorkList, "Result", gRow, colState
                strState = ""
                
                If strFe <> "" And strUIBC <> "" Then
                    strIntBase = "99"
                    strResult = Val(strFe) + Val(strUIBC)
                    
                    strFe = ""
                    strUIBC = ""
                    GoTo Rst
                End If
        End Select
    Next

    If strState = "Q" Then
        comEqp.Output = ENQ
        SetRawData "[Tx]" & ENQ
    End If


End Sub


''''-----------------------------------------------------------------------------'
''''   ???? : ?????κ? ?????? ?????? ????
''''-----------------------------------------------------------------------------'
'''Private Sub EditRcvDataAU()
'''    Dim strRcvBuf    As String   '?????? Data
'''    Dim strType      As String   '?????? Record Type
'''    Dim strBarno     As String   '?????? ???ڵ???ȣ
'''    Dim strSeq       As String   '?????? Sequence
'''    Dim strRackNo    As String   '?????? Rack Or Disk No
'''    Dim strTubePos   As String   '?????? Tube Position
'''    Dim strIntBase   As String   '?????? ???????? ?˻???
'''    Dim strResult    As String   '?????? ????
'''    Dim strQCResult  As String   '?????? ????(QC)
'''    Dim strFlag      As String   '?????? Abnormal Flag
'''    Dim strComm      As String   '?????? Comment
'''    Dim strTemp1     As String
'''    Dim strTemp2     As String
'''    Dim intCnt       As Integer
'''
'''    Dim lsExamCode As String
'''    Dim lsExamName As String
'''    Dim lsSeqNo As String
'''    Dim lsResult_Buff As String
'''    Dim lsExamDate As String
'''    Dim lsEquipRes As String
'''    Dim lsResRow    As String
'''    Dim ii As Integer
'''    Dim strTmp      As String
'''    Dim intIdx      As Integer
'''
'''
'''    For intCnt = 1 To UBound(strRecvData)
'''        strRcvBuf = strRecvData(intCnt)
'''        strType = Mid$(strRcvBuf, 1, 2)
'''
'''        Select Case strType
'''            '## Order Begin =========================================
'''            Case "RB"   '## Begin Inquiry Text
'''            Case "R "    '## Inquiry Order
'''                strBarno = Trim(Mid(strRcvBuf, 14, 20))
'''                strRackNo = Mid(strRcvBuf, 3, 4)
'''                strTubePos = Mid(strRcvBuf, 7, 2)
'''
'''                With mOrder
'''                    .BarNo = strBarno
'''                    .RackNo = strRackNo
'''                    .TubePos = strTubePos
'''                    .Seq = Mid(strRcvBuf, 9, 5)
'''                End With
'''
'''                Call GetOrder(strBarno)
'''
'''            Case "RE"   '## End Inquirty Text
'''
'''            '## Result =========================================
'''            Case "DB"   '## Begin Result Text
'''            Case "D "    '## Result
'''                strBarno = Trim$(Mid$(strRcvBuf, 14, 10))
'''
'''                With mResult
'''                    .BarNo = strBarno
'''                    .RackNo = Mid(strRcvBuf, 3, 4)
'''                    .TubePos = Mid(strRcvBuf, 7, 2)
'''                End With
'''
'''                If strBarno = "" Then Exit Sub
'''
'''                strTmp = Mid$(strRcvBuf, 29)
'''
'''                Call SetPatInfo(strBarno)
'''
'''                Do While Len(strTmp) >= 11
'''                    strIntBase = Mid$(strTmp, 2, 2)
'''                    strResult = Mid$(strTmp, 4, 6)
'''                    strComm = Mid$(strTmp, 10, 1)
'''
'''                    If strResult <> "" Then
'''                        SQL = ""
'''                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'''                        SQL = SQL & "  FROM EQPMASTER"
'''                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'''                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'''                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
'''
'''                        Res = GetDBSelectColumn(gLocal, SQL)
'''
'''                        '-- ???? ???? ????
'''                        If Res > 0 Then
'''                            lsExamCode = Trim(gReadBuf(0))
'''                            lsExamName = Trim(gReadBuf(1))
'''                            lsSeqNo = Trim(gReadBuf(2))
'''
'''                            lsResRow = vasRes.DataRowCnt + 1
'''                            If vasRes.MaxRows < lsResRow Then
'''                                vasRes.MaxRows = lsResRow
'''                            End If
'''
'''                            '?Ҽ??? ó??, ???? ???? ó??
'''                            lsEquipRes = strResult
'''                            strResult = SetResult(strResult, strIntBase)
'''                            lsResult_Buff = strResult
'''
'''                            '-- Work List
'''                            SetText vasID, "Result", gRow, colState                 '11 ????????
'''
'''
'''                            '-- ???? List
'''                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '?????ڵ?
'''                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '?˻??ڵ?
'''                            SetText vasRes, lsExamName, lsResRow, colExamName       '?˻???
'''                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '????????
'''                            SetText vasRes, strResult, lsResRow, colResult          '????
'''                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '????
'''                            SetText vasRes, strComm, lsResRow, 7                    'Flag
'''                            '-- ???? ????
'''                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                            lsResult_Buff = ""
'''
'''                            strState = "R"
'''
'''                        '-- ???? ???? ????
'''                        Else
'''
'''                                  SQL = "Select examcode, examname, seqno "
'''                            SQL = SQL & "  From EQPMASTER"
'''                            SQL = SQL & " Where equipno = '" & gEquip & "' "
'''                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
'''                            Res = GetDBSelectColumn(gLocal, SQL)
'''
'''                            If Res > 0 Then
'''                                lsExamCode = Trim(gReadBuf(0))
'''                                lsExamName = Trim(gReadBuf(1))
'''                                lsSeqNo = Trim(gReadBuf(2))
'''
'''                                lsResRow = vasRes.DataRowCnt + 1
'''                                If vasRes.MaxRows < lsResRow Then
'''                                    vasRes.MaxRows = lsResRow
'''                                End If
'''
'''                                '?Ҽ??? ó??, ???? ???? ó??
'''                                lsEquipRes = strResult
'''                                strResult = SetResult(strResult, strIntBase)
'''                                lsResult_Buff = strResult
'''
'''                                '-- Work List
'''                                SetText vasID, "Result", gRow, colState                 '????????
'''
'''                                '-- ???? List
'''                                SetText vasRes, strIntBase, lsResRow, colEquipCode      '?????ڵ?
'''                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '?˻??ڵ?
'''                                SetText vasRes, lsExamName, lsResRow, colExamName       '?˻???
'''                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '????????
'''                                SetText vasRes, strResult, lsResRow, colResult          '????
'''                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '????
'''                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
'''                                '-- ???? ????
'''                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
'''
'''                                lsResult_Buff = ""
'''
'''                            End If
'''                        End If
'''                    End If
'''                    strTmp = Mid$(strTmp, 12)
'''                Loop
'''
'''
'''                If MnTransAuto.Checked = True And strState = "R" Then
'''
'''                    Res = SaveTransDataW(gRow)
'''
'''                    If Res = -1 Then
'''                        '-- ???? ????
'''                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
'''                        SetText vasID, "Failed", gRow, colState
'''                    Else
'''                        '-- ???? ????
'''                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
'''                        SetText vasID, "Trans", gRow, colState
'''
'''                        SQL = " Update PATRESULT Set " & vbCrLf & _
'''                              " sendflag = '2' " & vbCrLf & _
'''                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''                              " And barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' "
'''                        Res = SendQuery(gLocal, SQL)
'''                        If Res = -1 Then
'''                            SaveQuery SQL
'''                            Exit Sub
'''                        End If
'''                    End If
'''                End If
'''
'''                'SetText vasID, "Result", gRow, colState
'''                strState = ""
'''
'''            Case "DE"   '## End Result Text
'''                strState = ""
'''        End Select
'''    Next
'''
'''End Sub

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
' asRow2 = ???? List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = Format(dtpToday, "yyyymmdd")

    SQL = ""
    SQL = "DELETE FROM PATRESULT " & vbCrLf & _
          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
          "  AND BARCODE = '" & Trim(GetText(vasWorkList, asRow1, colBarcode)) & "' " & vbCrLf & _
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
    'SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
    SQL = SQL & "'" & Trim(Format(dtpToday.Value, "YYYYMMDD")) & "', "
    SQL = SQL & "'" & gEquip & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colBarcode)) & "', "
    SQL = SQL & "'', "
    SQL = SQL & "'', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colPID)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colPName)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colSex)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colAge)) & "', "
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
    sMsg = "?˻??ڸ? ?Է????ּ???."
    lblUser.Caption = InputBox(sMsg, "?˻??? ?Է?")

End Sub





Private Sub txtNum_KeyPress(KeyAscii As Integer)
Dim intRow As Integer

    If KeyAscii = 13 Then
        With vasWorkList
            For intRow = .ActiveRow To .DataRowCnt
                '.Row = intRow
                '.Col = colCheckBox
                'If .Value = 1 Then
                    SetText vasWorkList, Format(txtRack, "00") & "-" & Format(txtNum, "00"), intRow, colSeqNo
    
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
    
    lsID = Trim(GetText(vasID, Row, colBarcode))
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    lblBarcode(0).Caption = lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPName))
    'Local???? ?ҷ?????
    ClearSpread vasRes
    
    '?????ڵ?, ?˻??ڵ?, ?˻???, ????, ????
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
'        If MsgBox("?ش? ȯ?ڰ????? ?????Ͻðڽ??ϱ??", vbInformation + vbYesNo, "?˸?") = vbNo Then
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
'        'Local???? ?ҷ?????
'        ClearSpread vasTemp
'
'        '?????ڵ?, ?˻??ڵ?, ?˻???, ????, ????
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
    'Local???? ?ҷ?????
    ClearSpread vasRRes
    
    '?????ڵ?, ?˻??ڵ?, ?˻???, ????, ????
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
'        'Local???? ?ҷ?????
'        ClearSpread vasTemp
'
'        '?????ڵ?, ?˻??ڵ?, ?˻???, ????, ????
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
'        If MsgBox("?ش? ȯ?ڰ????? ?????Ͻðڽ??ϱ??", vbInformation + vbYesNo, "?˸?") = vbNo Then
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
    'Local???? ?ҷ?????
    ClearSpread vasRes
    
    '?????ڵ?, ?˻??ڵ?, ?˻???, ????, ????
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

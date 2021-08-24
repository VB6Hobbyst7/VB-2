VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   "Hitachi7020 Interface "
   ClientHeight    =   10680
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   20670
   BeginProperty Font 
      Name            =   "±¼¸²Ã¼"
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
   ScaleHeight     =   10680
   ScaleWidth      =   20670
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   3375
      Left            =   15180
      TabIndex        =   47
      Top             =   6840
      Visible         =   0   'False
      Width           =   8625
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   240
         TabIndex        =   48
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
         Height          =   1275
         Left            =   240
         TabIndex        =   76
         Top             =   300
         Width           =   8055
         _Version        =   393216
         _ExtentX        =   14208
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
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmInterface.frx":2F98
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   6375
      Left            =   15150
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   960
         Left            =   1170
         TabIndex        =   18
         Top             =   5010
         Visible         =   0   'False
         Width           =   4155
         Begin MSCommLib.MSComm comEqp 
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
      Begin VB.CommandButton cmdRTrans 
         Caption         =   "°á°ú¼öµ¿Àü¼Û"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   435
         Left            =   7290
         TabIndex        =   59
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtTest 
         Height          =   675
         Left            =   4440
         TabIndex        =   58
         Top             =   1080
         Visible         =   0   'False
         Width           =   4125
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   405
         Left            =   5370
         TabIndex        =   55
         Top             =   2580
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1500
         Picture         =   "frmInterface.frx":75DA
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   38
         Top             =   5910
         Width           =   285
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   1545
         Left            =   240
         TabIndex        =   35
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
         SpreadDesigner  =   "frmInterface.frx":7B64
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   150
         TabIndex        =   33
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "±¼¸²"
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
         ScrollBars      =   2  '¼öÁ÷
         TabIndex        =   23
         Top             =   5220
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   4830
         TabIndex        =   22
         Top             =   5160
         Width           =   1125
      End
      Begin VB.TextBox Text_ini 
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4830
         TabIndex        =   21
         Top             =   5655
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1770
         MultiLine       =   -1  'True
         ScrollBars      =   3  '¾ç¹æÇâ
         TabIndex        =   20
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
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   19
         Top             =   5190
         Value           =   1  'È®ÀÎ
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   1755
         Left            =   3870
         TabIndex        =   17
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
         SpreadDesigner  =   "frmInterface.frx":7DA6
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1425
         Left            =   210
         TabIndex        =   24
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
         SpreadDesigner  =   "frmInterface.frx":7FE8
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   2175
         Left            =   3750
         TabIndex        =   25
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
         SpreadDesigner  =   "frmInterface.frx":822A
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   270
         TabIndex        =   60
         Top             =   1860
         Width           =   3375
         _Version        =   393216
         _ExtentX        =   5953
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
         SpreadDesigner  =   "frmInterface.frx":846C
      End
      Begin FPSpread.vaSpread vasExcel 
         Height          =   1185
         Left            =   6000
         TabIndex        =   65
         Top             =   5100
         Visible         =   0   'False
         Width           =   1935
         _Version        =   393216
         _ExtentX        =   3413
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
         SpreadDesigner  =   "frmInterface.frx":86AE
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'Åõ¸í
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   49
         Top             =   5910
         Width           =   1185
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3210
         TabIndex        =   27
         Top             =   5850
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   435
         Left            =   4050
         TabIndex        =   26
         Top             =   5820
         Width           =   915
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   10185
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   17965
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "WorkList"
      TabPicture(0)   =   "frmInterface.frx":88F0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "¹ÞÀº°á°ú"
      TabPicture(1)   =   "frmInterface.frx":890C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   9675
         Left            =   -74820
         TabIndex        =   9
         Top             =   360
         Width           =   14625
         Begin VB.ComboBox cboPartR 
            Height          =   315
            Left            =   2640
            TabIndex        =   78
            Text            =   "Combo1"
            Top             =   300
            Width           =   1005
         End
         Begin VB.CommandButton cmdRIFTrans 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6390
            TabIndex        =   77
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox chkQC 
            Caption         =   "QC¸¸Á¶È¸"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   10590
            TabIndex        =   67
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdXmlDel 
            Caption         =   "XMLÁ¤¸®"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11580
            TabIndex        =   61
            Top             =   270
            Width           =   1395
         End
         Begin VB.Frame Frame5 
            Height          =   825
            Left            =   10410
            TabIndex        =   28
            Top             =   690
            Width           =   4035
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   34
               Top             =   960
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   1
               Left            =   1500
               TabIndex        =   32
               Top             =   540
               Width           =   1305
            End
            Begin VB.Label Label4 
               Caption         =   "È¯ÀÚ¸í :"
               BeginProperty Font 
                  Name            =   "±¼¸²Ã¼"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   360
               TabIndex        =   31
               Top             =   540
               Width           =   1005
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   1
               Left            =   1515
               TabIndex        =   30
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label2 
               Caption         =   "Ã­Æ®¹øÈ£ :"
               BeginProperty Font 
                  Name            =   "±¼¸²Ã¼"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   150
               TabIndex        =   29
               Top             =   240
               Width           =   1230
            End
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "&Excel"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5040
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdRSch 
            Caption         =   "·ÎÄÃ°á°úÁ¶È¸"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3690
            TabIndex        =   14
            Top             =   240
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   13
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   125894657
            CurrentDate     =   40457
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   11
            Top             =   780
            Width           =   225
         End
         Begin VB.CommandButton cmdRClear 
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
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
            Top             =   270
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   7965
            Left            =   10410
            TabIndex        =   12
            Top             =   1605
            Width           =   4005
            _Version        =   393216
            _ExtentX        =   7064
            _ExtentY        =   14049
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²"
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
            SpreadDesigner  =   "frmInterface.frx":8928
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   8805
            Left            =   150
            TabIndex        =   46
            Top             =   750
            Width           =   10155
            _Version        =   393216
            _ExtentX        =   17912
            _ExtentY        =   15531
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":C6C7
            UserResize      =   2
         End
         Begin MSComCtl2.DTPicker dtpExamDate1 
            Height          =   315
            Left            =   9120
            TabIndex        =   63
            Top             =   -30
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   125894657
            CurrentDate     =   40457
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   8820
            TabIndex        =   62
            Top             =   0
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "°Ë»çÀÏÀÚ"
            BeginProperty Font 
               Name            =   "±¼¸²"
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
            TabIndex        =   45
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
         Begin VB.CheckBox chkWAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   900
            TabIndex        =   80
            Top             =   5490
            Width           =   225
         End
         Begin VB.CheckBox chkSave 
            Caption         =   "ÀúÀåÆ÷ÇÔ"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   7500
            TabIndex        =   79
            Top             =   150
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CommandButton cmdWUp 
            BackColor       =   &H00C0FFFF&
            Caption         =   "¡ã"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6630
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   74
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdWDown 
            BackColor       =   &H00C0FFFF&
            Caption         =   "¡å"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7020
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   73
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdDown 
            BackColor       =   &H00C0FFFF&
            Caption         =   "¡å"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5250
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   72
            Top             =   4950
            Width           =   525
         End
         Begin VB.CommandButton cmdUp 
            BackColor       =   &H00C0FFFF&
            Caption         =   "¡ã"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   71
            Top             =   4950
            Width           =   525
         End
         Begin VB.CommandButton cmdWPrint 
            Caption         =   "Ãâ·Â"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11250
            TabIndex        =   70
            Top             =   240
            Width           =   1005
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Ãâ·Â"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6420
            TabIndex        =   69
            Top             =   4950
            Width           =   1005
         End
         Begin VB.ComboBox cboPart 
            Height          =   315
            Left            =   3840
            TabIndex        =   68
            Text            =   "Combo1"
            Top             =   210
            Width           =   1005
         End
         Begin VB.TextBox txtNum 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   5850
            TabIndex        =   66
            Text            =   "0"
            Top             =   4980
            Width           =   435
         End
         Begin VB.CommandButton cmdDownLoad 
            Caption         =   "Down"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
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
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Á¶È¸"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4860
            TabIndex        =   50
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12330
            TabIndex        =   8
            Top             =   240
            Width           =   1005
         End
         Begin VB.CommandButton cmdIFTrans 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13410
            TabIndex        =   7
            Top             =   240
            Width           =   1005
         End
         Begin VB.Frame Frame6 
            Height          =   585
            Left            =   7530
            TabIndex        =   39
            Top             =   630
            Width           =   6945
            Begin VB.Label Label8 
               Caption         =   "Ã­Æ®¹øÈ£ :"
               BeginProperty Font 
                  Name            =   "±¼¸²Ã¼"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   44
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   0
               Left            =   1605
               TabIndex        =   43
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label6 
               Caption         =   "È¯ÀÚ¸í :"
               BeginProperty Font 
                  Name            =   "±¼¸²Ã¼"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3150
               TabIndex        =   42
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   0
               Left            =   4200
               TabIndex        =   41
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   40
               Top             =   720
               Width           =   1155
            End
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   900
            TabIndex        =   5
            Top             =   780
            Width           =   225
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8070
            Left            =   7530
            TabIndex        =   6
            Top             =   1425
            Width           =   6915
            _Version        =   393216
            _ExtentX        =   12197
            _ExtentY        =   14235
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²"
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
            SpreadDesigner  =   "frmInterface.frx":CF90
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
               Appearance      =   0  'Æò¸é
               Height          =   1455
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  '¼öÁ÷
               TabIndex        =   4
               Top             =   240
               Width           =   5775
            End
         End
         Begin MSComCtl2.DTPicker dtpToday 
            Height          =   315
            Left            =   8640
            TabIndex        =   36
            Top             =   270
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   125894656
            CurrentDate     =   40457
         End
         Begin FPSpread.vaSpread vasWorkList 
            Height          =   4155
            Left            =   180
            TabIndex        =   51
            Top             =   690
            Width           =   7215
            _Version        =   393216
            _ExtentX        =   12726
            _ExtentY        =   7329
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            MaxRows         =   20
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBarMaxAlign=   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":10CB4
            ScrollBarTrack  =   1
            ShowScrollTips  =   1
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   315
            Left            =   1080
            TabIndex        =   52
            Top             =   210
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   125894657
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   315
            Left            =   2520
            TabIndex        =   53
            Top             =   210
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   125894657
            CurrentDate     =   40457
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   4125
            Left            =   150
            TabIndex        =   75
            Top             =   5400
            Width           =   7215
            _Version        =   393216
            _ExtentX        =   12726
            _ExtentY        =   7276
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   16
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":1174A
            UserResize      =   2
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2370
            TabIndex        =   56
            Top             =   300
            Width           =   105
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "°Ë»çÀÏÀÚ"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   7620
            TabIndex        =   54
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblDate 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Á¶È¸ÀÏÀÚ"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   37
            Top             =   300
            Width           =   780
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '¾Æ·¡ ¸ÂÃã
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   10275
      Width           =   20670
      _ExtentX        =   36460
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13124
            MinWidth        =   13124
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "¿ÀÀü 9:45"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2019-09-18"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²Ã¼"
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
         Caption         =   "Á¾·á"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "Setting"
      Begin VB.Menu MnTConfig 
         Caption         =   "Åë½Å¼³Á¤"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "ÄÚµå¼³Á¤"
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
Const colSpecNo = 0 '¹Ì»ç¿ë
Const colCheckBox = 1
Const colBarcode = 2
Const colRack = 3
Const colDISK = 3
Const colPos = 4
Const colBarNo = 4
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
Const RS  As String = ""
Const GS  As String = ""
Const SB As String = ""  'Chr(11)
Const EB As String = ""   'Chr(28)

Dim strRecvData()   As String
Dim intPhase        As Integer
Dim strState        As String
Dim intBufCnt       As Integer
Dim blnIsETB        As Boolean
Dim intSndPhase     As Integer
Dim intFrameNo      As Integer

Dim blnSameRecord As Boolean

Private Type typeXMLInData
    Company     As String
    HospCode    As String
    ChartNo     As String
    PatName     As String
    PatJumin    As String
    PatNo       As String
    CommDate    As String
    ExamNo      As String
    ExamID      As String
    ComExamID   As String
    Specimen    As String
    Result      As String
    Reference   As String
    Remark      As String
    RsltDate    As String
    IOFlag      As String
End Type

Dim XMLInData As typeXMLInData

Private Type typeXMLOutData
    Company     As String
    HospCode    As String
    ChartNo     As String
    PatName     As String
    PatJumin    As String
    PatNo       As String
    CommDate    As String
    ExamNo      As String
    ExamID      As String
    ComExamID   As String
    Specimen    As String
    Result      As String
    Reference   As String
    Remark      As String
    RsltDate    As String
    IOFlag      As String
End Type

Dim XMLOutData As typeXMLOutData

Public RS1 As ADODB.Recordset
Const FieldCnt As Integer = 60
Const Fld_Div As String = ""
'===============================

Private Sub chkAll_Click()
    Dim iRow As Long
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasWorkList.DataRowCnt
            vasWorkList.Row = iRow
            vasWorkList.Col = 1
            
            vasWorkList.Value = 0
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

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, vasID.MaxCols, lRow, 1, lRow + 1
    SetActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
    
End Sub

Private Sub cmdDownLoad_Click()
    Dim intRow As Integer
    Dim j  As Integer
    
    j = 0
    With vasWorkList
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = colCheckBox
            If .Value = 1 Then
                Call vasWorkList_DblClick(1, intRow)
                
'                vasID.MaxRows = vasID.MaxRows + 1
'
'                SetText vasID, "1", vasID.MaxRows, 1
'                .Col = 2
'                SetText vasID, Trim(.Text), vasID.MaxRows, 2
'
'                .Col = 4
'                SetText vasID, Trim(.Text), vasID.MaxRows, 4
'                Call GetSampleInfoW(vasID.MaxRows)                                '5,6,7,8
                
                
                'Call .DeleteRows(intRow, intRow)
                '.MaxRows = .MaxRows - 1
                '.Action = ActionDeleteRow
'                .MaxRows = .MaxRows - 1

                txtNum = txtNum + 1
                
                .Col = 1
                .Value = "0"
                
            End If
        Next
        vasID.RowHeight(-1) = 12
    End With
End Sub

Private Sub cmdExcel_Click()
    Dim iRow As Integer
    Dim j As Integer
    Dim k As Integer
    
    
    Dim sCurDate As String
    Dim sSerDate As String
    Dim sHead As String
    Dim sFoot As String
    Dim sFileName As String
    
    Dim sA1c As String
    Dim sIFCC As String
    Dim seAg As String
    Dim blnWrite As Variant
    
    ClearSpread vasPrint

    blnWrite = False
    vasExcel.MaxRows = vasRID.MaxRows
    vasExcel.MaxCols = vasRID.MaxCols
    
    For iRow = 1 To vasRID.DataRowCnt
        vasRID.Row = iRow
        vasRID.Col = 1
            
        If vasRID.Value = 1 Then
            If blnWrite = False Then
                For j = 1 To vasRID.MaxCols
                    SetText vasExcel, Trim(GetText(vasRID, 0, j)), 0, j
                Next
            End If
            
            For j = 1 To vasRID.MaxCols
                SetText vasExcel, Trim(GetText(vasRID, iRow, j)), iRow, j
            Next
        End If
    Next iRow
    
    If vasExcel.DataRowCnt < 1 Then
        MsgBox "ÀúÀåÇÒ ÀÚ·á°¡ ¾ø½À´Ï´Ù.", , "¾Ë ¸²"
        Exit Sub
    Else
        CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All Files (*.*)|*.*"
        CommonDialog1.ShowSave
        sFileName = CommonDialog1.Filename
        SaveExcel sFileName, vasExcel
        MsgBox "¿¢¼¿ ÀúÀå¿Ï·á", vbOKOnly + vbInformation, Me.Caption
    End If
    
End Sub

Sub SaveExcel(Filename As String, argSpread As vaSpread)

On Error Resume Next

' Excel Object Library ¿Í ¿¬°áÇÕ´Ï´Ù.
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
'    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasWorkList.MaxRows = 0
'    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    
    
    For i = vasID.MaxRows To 1 Step -1
        If GetText(vasID, i, colState) = "Trans" Then
            Call DeleteRow(vasID, i, i)
            vasID.MaxRows = vasID.MaxRows - 1
        End If
    Next

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
                
                SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                      " TRANSYN = '2' " & vbCrLf & _
                      " WHERE EXAMTYPE = 'C' " & vbCrLf & _
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

Private Sub cmdPrint_Click()
    Dim iRow As Integer
    Dim j As Integer
    
    vasPrintBuf.MaxRows = 0
    'vasPrintBuf.MaxRows = 2
    
    With vasID
        For iRow = 1 To .DataRowCnt
            If iRow = 1 Then
                j = 1
                
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 2)), 0, 1:    vasPrintBuf.ColWidth(1) = 14
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 5)), 0, 2
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 6)), 0, 3
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 7)), 0, 4
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 8)), 0, 5
                SetText vasPrintBuf, Trim(GetText(vasID, 0, 9)), 0, 6
                SetText vasPrintBuf, "ºñ°í", 0, 7:                              vasPrintBuf.ColWidth(7) = 30
                
                vasPrintBuf.RowHeight(0) = 30
                
            End If
            
            
            If GetText(vasID, iRow, 1) = "1" Then
                vasPrintBuf.MaxRows = vasPrintBuf.MaxRows + 1
                
                SetText vasPrintBuf, Trim(GetText(vasID, iRow, 2)), vasPrintBuf.MaxRows, 1
                SetText vasPrintBuf, Trim(GetText(vasID, iRow, 5)), vasPrintBuf.MaxRows, 2
                SetText vasPrintBuf, Trim(GetText(vasID, iRow, 6)), vasPrintBuf.MaxRows, 3
                SetText vasPrintBuf, Trim(GetText(vasID, iRow, 7)), vasPrintBuf.MaxRows, 4
                SetText vasPrintBuf, Trim(GetText(vasID, iRow, 8)), vasPrintBuf.MaxRows, 5
                SetText vasPrintBuf, Trim(GetText(vasID, iRow, 9)), vasPrintBuf.MaxRows, 6
                SetText vasPrintBuf, " ", iRow, 7
                    
                vasPrintBuf.RowHeight(vasPrintBuf.MaxRows) = 23
            End If
            
        Next iRow
        
        'vasPrintbuf.RowHeight(-1) = 40

    End With
    
    If vasPrintBuf.DataRowCnt < 1 Then
        MsgBox "Ãâ·ÂÇÒ ÀÚ·á¸¦ ¼±ÅÃÇÏ¼¼¿ä", , "¾Ë ¸²"
        Exit Sub
    Else
        vasPrintBuf.PrintOrientation = PrintOrientationPortrait
        vasPrintBuf.Action = 13
    End If
    
    
'    vasWorkList.PrintOrientation = PrintOrientationPortrait '¼¼·ÎÃâ·Â
'    vasWorkList.Action = 13
    
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
    dtpExamDate1 = Date
    
End Sub

Private Sub cmdRIFTrans_Click()
    Dim lRow As Long
    
    If MsgBox("Á¶È¸µÈ °á°ú¸¦ Àü¼ÛÇÏ½Ã°Ú½À´Ï±î?", vbInformation + vbYesNo, "¾Ë¸²") = vbYes Then
            
        For lRow = 1 To vasRID.DataRowCnt
            vasRID.Row = lRow
            vasRID.Col = 1
            If vasRID.Value = 1 Then
                Res = SaveTransDataR(lRow)
            
                If Res = -1 Then
                    SetForeColor vasRID, lRow, lRow, 1, colState, 255, 0, 0
                    SetText vasRID, "Failed", lRow, colState
                Else
                    vasRID.Row = lRow
                    vasRID.Col = 1
                    vasRID.Value = 1
                    
                    SetBackColor vasRID, lRow, lRow, 1, colState, 202, 255, 112
                    SetText vasRID, "Trans", lRow, colState
                    
                    SQL = " UPDATE PAT_RES SET " & vbCrLf & _
                          " TRANSYN = '2' " & vbCrLf & _
                          " WHERE EXAMTYPE = 'C' " & vbCrLf & _
                          " AND BARCODE = '" & Trim(GetText(vasRID, lRow, 3)) & "' "
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
    End If
    
End Sub


Private Sub cmdRSch_Click()
    Dim RS1     As ADODB.Recordset
    Dim iRow    As Long
    Dim intCol  As Long
    Dim blnSame As Boolean
    Dim strDate As String
    Dim strChart As String
    Dim i As Integer
    
    ClearSpread vasRID
    ClearSpread vasRRes
    
    Call chkRAll_Click
    'Res = GetDBSelectVas(gLocal, SQL, vasRID)
    
          SQL = "SELECT DISTINCT COMMDATE, RSLTDATE, BARCODE,CHARTNO, PATNAME, PATSEX, PATAGE " & vbCrLf
    SQL = SQL & "  FROM PAT_RES " & vbCrLf
    SQL = SQL & " WHERE RSLTDATE BETWEEN '" & Format(dtpExamDate, "YYYYMMDD") & "' AND '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf
    'SQL = SQL & "   AND RESULT <> '' "
    SQL = SQL & "   AND EXAMTYPE = 'C' "
'    If chkQC.Value = "1" Then
'        SQL = SQL & "  AND PatName = 'QC' "
'    End If
    
    If cboPartR.ListIndex = 1 Then
        SQL = SQL & "  and mid(CHARTNO,1,1) <> 'G'  "
    ElseIf cboPartR.ListIndex = 2 Then
        SQL = SQL & "  and mid(CHARTNO,1,1) = 'G' "
    End If
    
    SQL = SQL & " GROUP BY COMMDATE, RSLTDATE, BARCODE,CHARTNO, PATNAME, PATSEX, PATAGE "
    
    cmdSQL.CommandText = SQL
    Set RS1 = cmdSQL.Execute
  
    If RS1.EOF = True Or RS1.BOF = True Then
        Exit Sub
    End If
    
    With vasRID
        While Not RS1.EOF
            For i = 1 To .DataRowCnt
                strDate = GetText(vasRID, i, 2)
                strChart = GetText(vasRID, i, 3)
                
                If Trim(RS1("RSLTDATE")) = strDate And Trim(RS1("CHARTNO")) = strChart Then
                    blnSame = True
                End If
                
'                If blnSame = True Then
'                    For intCol = 7 + 1 To vasID.MaxCols
'                        If Trim(RS1.Fields("EXAMID")) = gArrEquip(intCol - 7, 3) Then
'                            SetText vasRID, Trim(RS1.Fields("RESULT")) & "", .MaxRows, intCol
'                        End If
'                    Next
'                End If
            Next
            
            If blnSame = False Then
                iRow = iRow + 1
                .MaxRows = iRow
                .SetText colCheckBox, iRow, "1"
                .SetText 2, iRow, Trim(RS1.Fields("COMMDATE").Value) & ""
                .SetText 3, iRow, Trim(RS1.Fields("RSLTDATE").Value) & ""
                .SetText 4, iRow, Trim(RS1.Fields("BARCODE").Value) & ""
                .SetText 5, iRow, Trim(RS1.Fields("CHARTNO").Value) & ""
                .SetText 6, iRow, Trim(RS1.Fields("PATNAME").Value) & ""
                .SetText 7, iRow, Trim(RS1.Fields("PATSEX").Value) & ""
                .SetText 8, iRow, Trim(RS1.Fields("PATAGE").Value) & ""
                
'                For intCol = 7 + 1 To vasID.MaxCols
'                    If Trim(RS1.Fields("EXAMID")) = gArrEquip(intCol - 7, 3) Then
'                        SetText vasRID, Trim(RS1.Fields("RESULT")) & "", .MaxRows, intCol
'                    End If
'                Next
            End If
            
            blnSame = False
            RS1.MoveNext
        Wend
        .RowHeight(-1) = 12
    End With
    
    Set RS1 = Nothing
    
End Sub

Private Function getResultCnt(ByVal strBarNo As String) As String

    Dim RS1     As ADODB.Recordset
    
    getResultCnt = "0"
    
    
          SQL = "SELECT COUNT(*) as CNT " & vbCrLf
    SQL = SQL & "  FROM PAT_RES " & vbCrLf
    SQL = SQL & " WHERE TRANSDT = '" & Format(dtpExamDate, "YYYYMMDD") & "' " & vbCrLf
    SQL = SQL & "   AND RESULT <> '' "
    SQL = SQL & "   AND EXAMTYPE = 'C' "
    SQL = SQL & "   AND BARCODE = '" & strBarNo & "' "
    
    cmdSQL.CommandText = SQL
    Set RS1 = cmdSQL.Execute
  
    If RS1.EOF = True Or RS1.BOF = True Then
        Exit Function
    End If
    
    While Not RS1.EOF
        getResultCnt = Trim(RS1.Fields("CNT").Value) & ""
        RS1.MoveNext
    Wend

    Set RS1 = Nothing

End Function

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
                      "        TRANSYN = '2' " & vbCrLf & _
                      "  WHERE BARCODE = '" & Trim(GetText(vasRID, lRow, colBarcode)) & "' " & _
                      "    AND EXAMTYPE = 'C' "
                      
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


Private Function f_subSet_XMLWorkList(ByVal strDate As String, ByVal strDate1 As String, Optional ByVal strTime As String) As Variant
    Dim strPath   As String
    Dim strBuffer As String
    Dim i         As Long
    Dim lngBufLen As Long
    Dim BufChar   As String
    Dim strTmp As String
    Dim intIdx As Integer
    
    
On Error GoTo ErrorTrap
    
    Screen.MousePointer = 11
    
    '-- ¿À´õÆÄÀÏ¸í°ú °æ·Î¸¦ ÁöÁ¤ÇÑ´Ù.
    strPath = "C:\UBCare\SINAI\IF\ExamIF_In.xml"

    
    '1¶óÀÎ¾¿ °¡Á®¿À±â MSDN³»¿ë
    Dim TextLine
    Open strPath For Input As #1 ' ÆÄÀÏÀ» ¿±´Ï´Ù.
    
    Do While Not EOF(1) ' ÆÄÀÏÀÇ ³¡À» ¸¸³¯ ¶§±îÁö ¹Ýº¹ÇÕ´Ï´Ù.
        Line Input #1, TextLine ' º¯¼ö·Î µ¥ÀÌÅÍ ÇàÀ» ÀÐ¾îµéÀÔ´Ï´Ù.
        strBuffer = strBuffer & TextLine
    Loop
    
    Close #1 ' ÆÄÀÏÀ» ´Ý½À´Ï´Ù
 
    intIdx = 0
    lngBufLen = Len(strBuffer)
        
    For i = 1 To lngBufLen
        If intIdx = 0 Then
            BufChar = Mid$(strBuffer, i, 4)
        Else
            BufChar = Mid$(strBuffer, i + 3)
        End If
        
        If BufChar = "<°Ë»ç>" Then
            intIdx = 1
            strTmp = BufChar
        Else
            strTmp = strTmp & BufChar
            If intIdx = 1 Then Exit For
        End If
    
    Next
    
'    f_subSet_XMLWorkList = Split(strTmp, "</°Ë»ç>")
    strTmp = Replace(strTmp, "<°Ë»ç>", ""): strTmp = Replace(strTmp, "</°Ë»ç>", "|")
    strTmp = Replace(strTmp, "<¾÷Ã¼>", ""): strTmp = Replace(strTmp, "</¾÷Ã¼>", ",")
    strTmp = Replace(strTmp, "<¿ä¾ç±â°ü¹øÈ£>", ""): strTmp = Replace(strTmp, "</¿ä¾ç±â°ü¹øÈ£>", ",")
    strTmp = Replace(strTmp, "<Â÷Æ®¹øÈ£>", ""): strTmp = Replace(strTmp, "</Â÷Æ®¹øÈ£>", ",")
    strTmp = Replace(strTmp, "<¼öÁøÀÚ¸í>", ""): strTmp = Replace(strTmp, "</¼öÁøÀÚ¸í>", ",")
    strTmp = Replace(strTmp, "<ÁÖ¹Îµî·Ï¹øÈ£>", ""): strTmp = Replace(strTmp, "</ÁÖ¹Îµî·Ï¹øÈ£>", ",")
    strTmp = Replace(strTmp, "<³»¿ø¹øÈ£>", ""): strTmp = Replace(strTmp, "</³»¿ø¹øÈ£>", ",")
    strTmp = Replace(strTmp, "<ÀÇ·ÚÀÏ>", ""): strTmp = Replace(strTmp, "</ÀÇ·ÚÀÏ>", ",")
    strTmp = Replace(strTmp, "<°Ë»ç¹øÈ£>", ""): strTmp = Replace(strTmp, "</°Ë»ç¹øÈ£>", ",")
    strTmp = Replace(strTmp, "<°Ë»çID>", ""): strTmp = Replace(strTmp, "</°Ë»çID>", ",")
    strTmp = Replace(strTmp, "<¾÷Ã¼°Ë»çID>", ""): strTmp = Replace(strTmp, "</¾÷Ã¼°Ë»çID>", ",")
    strTmp = Replace(strTmp, "<°ËÃ¼>", ""): strTmp = Replace(strTmp, "</°ËÃ¼>", ",")
    strTmp = Replace(strTmp, "<°á°úÄ¡>", ""): strTmp = Replace(strTmp, "</°á°úÄ¡>", ",")
    strTmp = Replace(strTmp, "<ÂüÁ¶Ä¡>", ""): strTmp = Replace(strTmp, "</ÂüÁ¶Ä¡>", ",")
    strTmp = Replace(strTmp, "<¼Ò°ß>", ""): strTmp = Replace(strTmp, "</¼Ò°ß>", ",")
    strTmp = Replace(strTmp, "<°á°úÀÏ>", ""): strTmp = Replace(strTmp, "</°á°úÀÏ>", ",")
    strTmp = Replace(strTmp, "<¾÷Ã¼>", ""): strTmp = Replace(strTmp, "</¾÷Ã¼>", ",")
    strTmp = Replace(strTmp, "<ÀÔ¿ø¿Ü·¡±¸ºÐ>", ""): strTmp = Replace(strTmp, "</ÀÔ¿ø¿Ü·¡±¸ºÐ>", ",")
    
    f_subSet_XMLWorkList = Split(strTmp, "|")
    blnSameRecord = True
    
    
    Call SetSQLData("¿À´õÀúÀå", strTmp, "A")
    Kill strPath
    
    Screen.MousePointer = 0

    
    Exit Function
        
ErrorTrap:
    
    blnSameRecord = False
    Screen.MousePointer = 0
    
    
End Function


Private Function SeqSearch_New(ByVal brspread As Object, ByVal brSeq As String, ByVal brSeq2 As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch_New = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = brSeq Then
                .Row = sCnt
                .Col = 3
                If Trim(.Text) = brSeq2 Then
                    SeqSearch_New = sCnt
                    .Action = ActionActiveCell
                    .Refresh
                    Exit For
                End If
            End If
        Next sCnt
    End With

End Function


Private Function SeqSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = brSeq Then
                .Row = sCnt
                .Col = 5
                SeqSearch = .Row
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Function SeqNullSearch(ByVal brspread As Object, ByVal brSeq As String, ByVal brCol As Integer) As Long
Dim sCnt As Long

    SeqNullSearch = 0
    If brspread.MaxRows <= 0 Then
        Exit Function
    End If
    
    With brspread
        For sCnt = 1 To .MaxRows
            .Row = sCnt
            .Col = brCol
            If Trim(.Text) = "" Then
                SeqNullSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

End Function

Private Sub cmdSearch_Click()
'''    Dim sSch1, sSch2 As String
'''    Dim iRow As Integer
'''    Dim i, X As Long
'''    Dim sCnt As String
'''    Dim sExamCode As String
'''    Dim sExamName As String
'''    Dim FilNum
'''    Dim TxtString As String
'''    Dim TxtRece As String
'''    Dim PChartNum As String
'''    Dim PNAME As String
'''    Dim PJumin As String
'''    Dim PID As String
'''    Dim PExamCode As String
'''    Dim PReceDate As String
'''    Dim PAGE As String
'''    Dim PSEX As String
'''    Dim STxt, NumTxt As Long
'''    Dim SQL As String
'''    Dim PEquipno As String
'''    Dim PExamname As String
'''    Dim PEquipCode As String
'''    Dim pEqipType  As String
'''    Dim j As Long
'''    Dim BarFlag As Integer
'''    Dim TxtPat As String
'''    Dim TestNum, IOGubun As String
'''    Dim FindFile As String
'''    Dim StartDate As String
'''    Dim EndDate As String
'''    Dim varXML      As Variant
'''    Dim varTmp      As Variant
'''    Dim strBarNo As String
'''    Dim intCnt As Integer
'''    Dim pGrid_Point As Integer
'''    Dim sList As Integer
'''    Dim strBarNum As String
'''    Dim strSrcfile  As String
'''    Dim strDestFile As String
'''    Dim RSX  As ADODB.Recordset
'''    Dim InsNo   As String
'''
'''    Screen.MousePointer = 11
'''    vasWorkList.ReDraw = False
'''
'''    ClearSpread vasWorkList
'''    vasWorkList.MaxRows = 0
'''
'''          SQL = "select distinct commdate,chartno,patname,patsex,patage,remark from pat_res "
'''    SQL = SQL & " where commdate between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' AND '" & Format(dtpStopDt.Value, "yyyymmdd") & "'"
'''    If cboPart.ListIndex = 1 Then
'''        SQL = SQL & "  and mid(chartno,1,1) <> 'G'  "
'''    ElseIf cboPart.ListIndex = 2 Then
'''        SQL = SQL & "  and mid(chartno,1,1) = 'G' "
'''    'ElseIf cboPart.ListIndex = 3 Then
'''    '    SQL = SQL & "  and mid(chartno,1,1) = 'C' "
'''    End If
'''
''''        SQL = SQL & "   and (result = '' or result is null)"
'''
'''    If chkSave.Value = "0" Then
'''        SQL = SQL & "   and result = '' "
'''    End If
'''
'''    SQL = SQL & " Order by commdate,remark "
'''
'''    Set RSX = cn.Execute(SQL)
'''    Do Until RSX.EOF
'''        With vasWorkList
'''            'pGrid_Point = SeqSearch(vasWorkList, Trim(RSX.Fields("CHARTNO")), 5)
'''            pGrid_Point = SeqSearch(vasWorkList, Trim(RSX.Fields("CHARTNO")), 5)
'''
'''            If pGrid_Point = 0 Then
'''                pGrid_Point = SeqNullSearch(vasWorkList, Trim(RSX.Fields("CHARTNO")), 5)
'''                If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
'''                .RowHeight(-1) = 12
'''            End If
'''
'''            .SetText 1, pGrid_Point, "1"
'''            .SetText 2, pGrid_Point, Format(Trim(RSX.Fields("COMMDATE")), "####-##-##")
'''            .SetText 3, pGrid_Point, "C"
'''            strBarNum = Mid(Format(Trim(RSX.Fields("COMMDATE")), "########"), 5, 4) & Format(Trim(RSX.Fields("CHARTNO")), "0000000000")
'''            .SetText 4, pGrid_Point, strBarNum
'''            .SetText 5, pGrid_Point, Trim(RSX.Fields("CHARTNO"))
'''            .SetText 6, pGrid_Point, Trim(RSX.Fields("PATNAME"))
'''            .SetText 7, pGrid_Point, Trim(RSX.Fields("PATSEX"))
'''            .SetText 8, pGrid_Point, Trim(RSX.Fields("PATAGE"))
'''            .SetText 9, pGrid_Point, "Order"
'''
'''        End With
'''        RSX.MoveNext
'''    Loop
'''    RSX.Close
'''
'''    '-- XML ÀÏ±â
'''    varXML = f_subSet_XMLWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'''
'''
'''    If blnSameRecord = False Then
'''        'MsgBox "°Ë»ç ´ë»óÀÚ°¡ ¾ø½À´Ï´Ù.", vbOKOnly + vbInformation, App.Title
'''        Exit Sub
'''    End If
'''
'''    If UBound(varXML) < 1 Then
'''        'MsgBox "°Ë»ç ´ë»óÀÚ°¡ ¾ø½À´Ï´Ù.", vbOKOnly + vbInformation, App.Title
'''        Exit Sub
'''    Else
'''        strBarNo = ""
'''
'''        With vasWorkList
'''            For intCnt = 0 To UBound(varXML) - 1
'''                varTmp = Split(varXML(intCnt), ",")
'''
'''                '-- ÀåºñÃ¤³Î°ªÃ£±â
'''                SQL = ""
'''                SQL = SQL & " SELECT EQUIPCODE "
'''                SQL = SQL & "   FROM EQUIPEXAM"
'''                SQL = SQL & "  WHERE EXAMCODE = '" & Trim(varTmp(8)) & "' "
'''
'''                Res = GetDBSelectColumn(gLocal, SQL)
'''                XMLInData.ComExamID = ""
'''
'''                '-- ¿À´õ ÀÖÀ» °æ¿ì
'''                If Res > 0 Then
'''
'''                    XMLInData.ComExamID = Trim(gReadBuf(0))
'''
'''                    XMLInData.Company = varTmp(0)
'''                    XMLInData.HospCode = varTmp(1)
'''                    XMLInData.ChartNo = varTmp(2)
'''                    XMLInData.PatName = varTmp(3)
'''                    XMLInData.PatJumin = varTmp(4)
'''                    XMLInData.PatNo = varTmp(5)
'''                    XMLInData.CommDate = varTmp(6)
'''                    XMLInData.ExamNo = varTmp(7)
'''                    XMLInData.ExamID = varTmp(8)
'''                    'XMLInData.ComExamID = varTmp(9)
'''                    XMLInData.Specimen = varTmp(10)
'''                    XMLInData.Result = varTmp(11)
'''                    XMLInData.Reference = varTmp(12)
'''                    XMLInData.Remark = varTmp(13)
'''                    XMLInData.RsltDate = varTmp(14)
'''                    XMLInData.IOFlag = varTmp(15)
'''
'''                    SQL = ""
'''                    SQL = SQL & "select equipno, equipcode, examname, examtype "
'''                    SQL = SQL & "  from equipexam "
'''                    SQL = SQL & " where examcode = '" & XMLInData.ExamID & "' "
'''                    Res = db_select_Col(gLocal, SQL)
'''                    If Res > 0 Then
'''                        PEquipno = gReadBuf(0)
'''                        PEquipCode = gReadBuf(1)
'''                        PExamname = gReadBuf(2)
'''
'''                        If strBarNo <> XMLInData.ChartNo Or pEqipType <> gReadBuf(3) Then
'''                            pEqipType = gReadBuf(3)
'''
'''                            pGrid_Point = SeqSearch(vasWorkList, XMLInData.ChartNo, 5)
'''
'''                            If pGrid_Point = 0 Then
'''                                pGrid_Point = SeqNullSearch(vasWorkList, XMLInData.ChartNo, 5)
'''                                If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
'''                                .RowHeight(-1) = 12
'''                            End If
'''
'''                            .SetText 1, pGrid_Point, "1"
'''                            .SetText 2, pGrid_Point, Format(XMLInData.CommDate, "####-##-##")
'''                            .SetText 3, pGrid_Point, pEqipType
'''                            strBarNum = Mid(XMLInData.CommDate, 5, 4) & Format(XMLInData.ChartNo, "0000000000")
'''                            .SetText 4, pGrid_Point, strBarNum
'''                            .SetText 5, pGrid_Point, XMLInData.ChartNo
'''                            .SetText 6, pGrid_Point, XMLInData.PatName
'''                                        PJumin = Left(XMLInData.PatJumin, 6) & Right(XMLInData.PatJumin, 7)
'''                                        Call CalAgeSex(PJumin, Format(Date, "yyyy/mm/dd"))
'''                            .SetText 7, pGrid_Point, gPatGen.Sex
'''                            .SetText 8, pGrid_Point, gPatGen.Age
'''                            .SetText 9, pGrid_Point, "Order"
'''
'''                            InsNo = getMaxTestNum(XMLInData.CommDate)
'''
'''                        End If
'''                              SQL = "Select ChartNo from pat_res "
'''                        SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
'''                        SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
'''                        SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
'''                        SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
'''                        SQL = SQL & "   and ExamType = '" & pEqipType & "'"
'''
'''                        Res = db_select_Col(gLocal, SQL)
'''
'''                        If Res = 0 Then
'''                                  SQL = " insert into pat_res("
'''                            SQL = SQL & "Company,HospCode,ChartNo, "
'''                            SQL = SQL & "PatName,PatSex,PatAge,PatJumin,PatNo,"
'''                            SQL = SQL & "CommDate,ExamNo,ExamID,ComExamID, "
'''                            SQL = SQL & "Specimen,Result,Reference,Remark,RsltDate,IOFlag,BarCode,ExamType)"
'''                            SQL = SQL & " values ("
'''                            SQL = SQL & "'" & XMLInData.Company & "',"
'''                            SQL = SQL & "'" & XMLInData.HospCode & "',"
'''                            SQL = SQL & "'" & XMLInData.ChartNo & "',"
'''                            SQL = SQL & "'" & XMLInData.PatName & "',"
'''                            SQL = SQL & "'" & gPatGen.Sex & "',"
'''                            SQL = SQL & "'" & gPatGen.Age & "',"
'''                            SQL = SQL & "'" & XMLInData.PatJumin & "',"
'''                            SQL = SQL & "'" & XMLInData.PatNo & "',"
'''                            SQL = SQL & "'" & XMLInData.CommDate & "',"
'''                            SQL = SQL & "'" & XMLInData.ExamNo & "',"
'''                            SQL = SQL & "'" & XMLInData.ExamID & "',"
'''                            SQL = SQL & "'" & XMLInData.ComExamID & "',"
'''                            SQL = SQL & "'" & XMLInData.Specimen & "',"
'''                            SQL = SQL & "'" & XMLInData.Result & "',"
'''                            SQL = SQL & "'" & XMLInData.Reference & "',"
'''                            'SQL = SQL & "'" & XMLInData.Remark & "',"
'''                            SQL = SQL & "'" & InsNo & "',"
'''                            SQL = SQL & "'" & XMLInData.RsltDate & "',"
'''                            SQL = SQL & "'" & XMLInData.IOFlag & "',"
'''                            SQL = SQL & "'" & strBarNum & "',"
'''                            SQL = SQL & "'" & pEqipType & "')"
'''
'''                            Res = SendQuery(gLocal, SQL)
'''
'''                            If Res = -1 Then
'''                                SaveQuery SQL
'''                            End If
'''
'''                        '-- ¼ÓµµÇâ»óÀ» À§ÇØ Äõ¸®¹® Áö¿ì±â
'''                        Else
'''                                  SQL = " Update pat_res Set "
'''                            SQL = SQL & " PatName = '" & XMLInData.PatName & "', "
'''                            SQL = SQL & " PatSex  = '" & gPatGen.Sex & "' "
'''                            SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
'''                            SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
'''                            SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
'''                            SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
'''                            SQL = SQL & "   and ExamType = '" & pEqipType & "'"
'''
'''                            Res = SendQuery(gLocal, SQL)
'''                        End If
'''
'''                        strBarNo = XMLInData.ChartNo
'''                    End If
'''                End If
'''
'''                XMLInData.ComExamID = ""
'''            Next
'''
'''        End With
'''    End If
'''
'''    If chkSave.Value = "0" Then
'''        Call SaveCheck
'''    End If
'''
'''    vasWorkList.ReDraw = True
'''    Screen.MousePointer = 0


'-- 2017.12.19
    Dim sSch1, sSch2 As String
    Dim iRow As Integer
    Dim i, X As Long
    Dim sCnt As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim FilNum
    Dim TxtString As String
    Dim TxtRece As String
    Dim PChartNum As String
    Dim PNAME As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAGE As String
    Dim PSEX As String
    Dim STxt, NumTxt As Long
    Dim SQL As String
    Dim PEquipno As String
    Dim PExamname As String
    Dim PEquipCode As String
    Dim pEqipType  As String
    Dim j As Long
    Dim BarFlag As Integer
    Dim TxtPat As String
    Dim TestNum, IOGubun As String
    Dim FindFile As String
    Dim StartDate As String
    Dim EndDate As String
    Dim varXML      As Variant
    Dim varTmp      As Variant
    Dim strBarNo As String
    Dim intCnt As Integer
    Dim pGrid_Point As Integer
    Dim sList As Integer
    Dim strBarNum As String
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim RSX  As ADODB.Recordset
    Dim InsNo   As String
    
    Screen.MousePointer = 11
    
    vasRes.MaxRows = 0
    
    '-- XML ÀÏ±â
    varXML = f_subSet_XMLWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    
    If blnSameRecord = False Then
        GoTo Rst
    End If
    
    pEqipType = "C"
    
    If UBound(varXML) > 1 Then
        strBarNo = ""

        With vasWorkList
            For intCnt = 0 To UBound(varXML) - 1
                
                varTmp = Split(varXML(intCnt), ",")
                                
                '-- ÀåºñÃ¤³Î°ªÃ£±â
                SQL = ""
                SQL = SQL & " SELECT EQUIPCODE "
                SQL = SQL & "   FROM EQUIPEXAM"
                SQL = SQL & "  WHERE EXAMCODE = '" & Trim(varTmp(8)) & "' "
                
                Res = GetDBSelectColumn(gLocal, SQL)
                XMLInData.ComExamID = ""
                
                '-- ¿À´õ ÀÖÀ» °æ¿ì
                If Res > 0 Then
                    
                    If strBarNo <> XMLInData.ChartNo Then
                        InsNo = getMaxTestNum(XMLInData.CommDate)
                    End If
                    
                    XMLInData.ComExamID = Trim(gReadBuf(0))
                    XMLInData.Company = varTmp(0)
                    XMLInData.HospCode = varTmp(1)
                    XMLInData.ChartNo = varTmp(2)
                    XMLInData.PatName = varTmp(3)
                    XMLInData.PatJumin = varTmp(4)
                    XMLInData.PatNo = varTmp(5)
                    XMLInData.CommDate = varTmp(6)
                    XMLInData.ExamNo = varTmp(7)
                    XMLInData.ExamID = varTmp(8)
                    'XMLInData.ComExamID = varTmp(9)
                    XMLInData.Specimen = varTmp(10)
                    XMLInData.Result = varTmp(11)
                    XMLInData.Reference = varTmp(12)
                    XMLInData.Remark = varTmp(13)
                    XMLInData.RsltDate = varTmp(14)
                    XMLInData.IOFlag = varTmp(15)
                
                    strBarNum = Mid(Format(XMLInData.CommDate, "########"), 5, 4) & Format(XMLInData.ChartNo, "0000000000")
                    
                    PJumin = Left(XMLInData.PatJumin, 6) & Right(XMLInData.PatJumin, 7)
                    Call CalAgeSex(PJumin, Format(Date, "yyyy/mm/dd"))
                
                    'InsNo = getMaxTestNum(XMLInData.CommDate)
                          
                          SQL = "Select ChartNo from pat_res "
                    SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
                    SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
                    SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
                    SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
                    SQL = SQL & "   and ExamType = '" & pEqipType & "'"
                    
                    Res = db_select_Col(gLocal, SQL)
                    
                    If Res = 0 Then
                              SQL = " insert into pat_res("
                        SQL = SQL & "Company,HospCode,ChartNo, "
                        SQL = SQL & "PatName,PatSex,PatAge,PatJumin,PatNo,"
                        SQL = SQL & "CommDate,ExamNo,ExamID,ComExamID, "
                        SQL = SQL & "Specimen,Result,Reference,Remark,RsltDate,IOFlag,BarCode,ExamType)"
                        SQL = SQL & " values ("
                        SQL = SQL & "'" & XMLInData.Company & "',"
                        SQL = SQL & "'" & XMLInData.HospCode & "',"
                        SQL = SQL & "'" & XMLInData.ChartNo & "',"
                        SQL = SQL & "'" & XMLInData.PatName & "',"
                        SQL = SQL & "'" & gPatGen.Sex & "',"
                        SQL = SQL & "'" & gPatGen.Age & "',"
                        SQL = SQL & "'" & XMLInData.PatJumin & "',"
                        SQL = SQL & "'" & XMLInData.PatNo & "',"
                        SQL = SQL & "'" & XMLInData.CommDate & "',"
                        SQL = SQL & "'" & XMLInData.ExamNo & "',"
                        SQL = SQL & "'" & XMLInData.ExamID & "',"
                        SQL = SQL & "'" & XMLInData.ComExamID & "',"
                        SQL = SQL & "'" & XMLInData.Specimen & "',"
                        SQL = SQL & "'" & XMLInData.Result & "',"
                        SQL = SQL & "'" & XMLInData.Reference & "',"
                        'SQL = SQL & "'" & XMLInData.Remark & "',"
                        SQL = SQL & "'" & InsNo & "',"
                        SQL = SQL & "'" & XMLInData.RsltDate & "',"
                        SQL = SQL & "'" & XMLInData.IOFlag & "',"
                        SQL = SQL & "'" & strBarNum & "',"
                        SQL = SQL & "'" & pEqipType & "')"
                        
                        Res = SendQuery(gLocal, SQL)
                        
                        If Res = -1 Then
                            SaveQuery SQL
                        End If
                    
                    Else
                              SQL = " Update pat_res Set "
                        SQL = SQL & " PatName = '" & XMLInData.PatName & "', "
                        SQL = SQL & " PatSex  = '" & gPatGen.Sex & "' "
                        SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
                        SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
                        SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
                        SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
                        SQL = SQL & "   and ExamType = '" & pEqipType & "'"
                        
                        Res = SendQuery(gLocal, SQL)
                    End If
                    
                    strBarNo = XMLInData.ChartNo
                End If
                
                XMLInData.ComExamID = ""
            Next
            
        End With
    End If

Rst:
    
    vasWorkList.ReDraw = False
    
    ClearSpread vasWorkList
    vasWorkList.MaxRows = 0

          SQL = "select distinct commdate,chartno,patname,patsex,patage,remark from pat_res "
    SQL = SQL & " where commdate between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' AND '" & Format(dtpStopDt.Value, "yyyymmdd") & "'"
    If cboPart.ListIndex = 1 Then
        SQL = SQL & "  and mid(chartno,1,1) <> 'G'  "
    ElseIf cboPart.ListIndex = 2 Then
        SQL = SQL & "  and mid(chartno,1,1) = 'G' "
    End If
    
    If chkSave.Value = "0" Then
        SQL = SQL & "   and result = '' "
    End If
    
    SQL = SQL & " Order by commdate,remark "
    
    Set RSX = cn.Execute(SQL)
    Do Until RSX.EOF
        With vasWorkList
            pGrid_Point = SeqSearch(vasWorkList, Trim(RSX.Fields("CHARTNO")), 5)

            If pGrid_Point = 0 Then
                pGrid_Point = SeqNullSearch(vasWorkList, Trim(RSX.Fields("CHARTNO")), 5)
                If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
                .RowHeight(-1) = 12
            End If
            
            .SetText 1, pGrid_Point, "1"
            .SetText 2, pGrid_Point, Format(Trim(RSX.Fields("COMMDATE")), "####-##-##")
            .SetText 3, pGrid_Point, "C"
            strBarNum = Mid(Format(Trim(RSX.Fields("COMMDATE")), "########"), 5, 4) & Format(Trim(RSX.Fields("CHARTNO")), "0000000000")
            .SetText 4, pGrid_Point, strBarNum
            .SetText 5, pGrid_Point, Trim(RSX.Fields("CHARTNO"))
            .SetText 6, pGrid_Point, Trim(RSX.Fields("PATNAME"))
            .SetText 7, pGrid_Point, Trim(RSX.Fields("PATSEX"))
            .SetText 8, pGrid_Point, Trim(RSX.Fields("PATAGE"))
            .SetText 9, pGrid_Point, "Order"
        
        End With
        RSX.MoveNext
    Loop
    RSX.Close
        
        
    If chkSave.Value = "0" Then
        Call SaveCheck
    End If
    
    vasWorkList.ReDraw = True
    Screen.MousePointer = 0
    
End Sub

'-- ¿À´Ã °Ë»çÇÑ ³¯Â¥ÀÇ Max + 1 ¹øÈ£¸¦ °¡Á®¿Â´Ù
Private Function getMaxTestNum(ByVal strDate As String) As Long

    getMaxTestNum = 1
    
    '-- °á°ú¾÷µ¥ÀÌÆ®
          SQL = "SELECT MAX(REMARK) as SEQ FROM PAT_RES  "
    SQL = SQL & " WHERE MID(COMMDATE,1,8) = '" & strDate & "' " & vbCrLf
    
    Res = GetDBSelectColumn(gLocal, SQL)
    
    If Res > 0 Then
        If Trim(gReadBuf(0)) = "" Then
            getMaxTestNum = 1
        Else
            getMaxTestNum = Trim(gReadBuf(0)) + 1
        End If
    End If
    
    If getMaxTestNum >= 99999 Then
        getMaxTestNum = 99999
    End If
    
End Function


Private Sub SaveCheck()
    Dim i           As Integer
    Dim strDate     As String
    Dim strChart    As String
    Dim RS1         As ADODB.Recordset
    
    For i = 1 To vasWorkList.MaxRows
        strDate = Format(GetText(vasWorkList, i, 2), "yyyymmdd")
        strChart = GetText(vasWorkList, i, 5)
        
        '-- ±âÁ¸¿¡ ÀÖ´ÂÁö Á¶È¸
              SQL = "Select Count(*) from pat_res "
        SQL = SQL & " Where ChartNo  = '" & strChart & "' "
        SQL = SQL & "   and CommDate = '" & strDate & "'"
        'SQL = SQL & "   AND RSLTDATE BETWEEN '" & Format(dtpExamDate, "YYYYMMDD") & "' AND '" & Format(dtpExamDate1, "YYYYMMDD") & "' " & vbCrLf
        SQL = SQL & "   and RESULT = '' "
        
        cmdSQL.CommandText = SQL
        Set RS1 = cmdSQL.Execute
               
        If Not (RS1.EOF Or RS1.BOF) Then
            'rs.MoveFirst
        End If
        
        Do While Not RS1.EOF
            If RS1.Fields.Item(0).Value = 0 Then
                Call vasWorkList.DeleteRows(i, i)
                vasWorkList.MaxRows = vasWorkList.MaxRows - 1
            Else
                Exit Do
            End If
            RS1.MoveNext
            Exit Do
        Loop
        RS1.Close
    Next
    
End Sub

Public Function db_select_Col(argServer As Integer, argSQL As String) As Integer
'Äõ¸® ½ÇÇà ³»¿ëÀ» gReadbuf()ÀÇ Array¿¡ ÀúÀå
    Dim i, j As Integer
    
On Error GoTo ErrHandle
       
    db_select_Col = -1
    i = 0
    
    gReadBuf(0) = ""
    gReadBuf(1) = ""
    
    Select Case argServer
    Case gServer
        Set cmdSQL.ActiveConnection = cn_Ser
'    Case gServer1
'        Set cmdSQL.ActiveConnection = cn_Ser1
    Case gLocal
        Set cmdSQL.ActiveConnection = cn
    Case Else
        Exit Function
    End Select
    cmdSQL.CommandText = argSQL
    Set RS1 = cmdSQL.Execute
           
    If Not (RS1.EOF Or RS1.BOF) Then
        'rs.MoveFirst
    Else
        db_select_Col = 0
        gReadBuf(0) = ""
        RS1.Close
        Exit Function
    End If
    
    
    Do While Not RS1.EOF
        For i = 0 To RS1.Fields.Count - 1
            If IsNull(RS1.Fields.Item(i).Value) = True Then
                gReadBuf(i) = ""
            Else
                gReadBuf(i) = Trim(CStr(RS1.Fields.Item(i).Value))
            End If
        Next i
        
        db_select_Col = 1
        
        RS1.MoveNext
        Exit Do
    Loop
    
    RS1.Close
    
    Exit Function
    
ErrHandle:
    ''MsgBox Err.Number & " : " & Error(Err.Number), vbCritical
    db_select_Col = -1
End Function

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, vasID.MaxCols, lRow, 1, lRow - 1
    SetActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1

End Sub

Public Function SetActiveCell(ByRef vasTable As Object, ByVal vasRow As Integer, ByVal vasCol As Integer) As Boolean
'Æ¯Á¤ Cell ÁöÁ¤
    vasTable.Row = vasRow
    vasTable.Col = vasCol
    vasTable.Action = 0
End Function


Private Sub cmdWDown_Click()
    Dim lRow As Long
    
    lRow = vasWorkList.ActiveRow
    
    vasWorkList.SwapRange 1, lRow, vasWorkList.MaxCols, lRow, 1, lRow + 1
    SetActiveCell vasWorkList, lRow + 1, 2
'    vasWorkList_Click 2, lRow + 1
    
End Sub

Private Sub cmdWPrint_Click()
    Dim iRow As Integer
    Dim j As Integer
    
    vasPrintBuf.MaxRows = 0
    
    With vasWorkList
        For iRow = 1 To .DataRowCnt
            If iRow = 1 Then
                j = 1
                
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 2)), 0, 1:    vasPrintBuf.ColWidth(1) = 14
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 5)), 0, 2
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 6)), 0, 3
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 7)), 0, 4
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 8)), 0, 5
                SetText vasPrintBuf, Trim(GetText(vasWorkList, 0, 9)), 0, 6
                SetText vasPrintBuf, "ºñ°í", 0, 7:                              vasPrintBuf.ColWidth(7) = 30
                
                vasPrintBuf.RowHeight(0) = 30
                
            End If
            
            
            If GetText(vasWorkList, iRow, 1) = "1" Then
                vasPrintBuf.MaxRows = vasPrintBuf.MaxRows + 1
                
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 2)), vasPrintBuf.MaxRows, 1
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 5)), vasPrintBuf.MaxRows, 2
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 6)), vasPrintBuf.MaxRows, 3
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 7)), vasPrintBuf.MaxRows, 4
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 8)), vasPrintBuf.MaxRows, 5
                SetText vasPrintBuf, Trim(GetText(vasWorkList, iRow, 9)), vasPrintBuf.MaxRows, 6
                SetText vasPrintBuf, " ", iRow, 7
                    
                vasPrintBuf.RowHeight(vasPrintBuf.MaxRows) = 23
            End If
            
        Next iRow
        
    End With
    
    If vasPrintBuf.DataRowCnt < 1 Then
        MsgBox "Ãâ·ÂÇÒ ÀÚ·á¸¦ ¼±ÅÃÇÏ¼¼¿ä", , "¾Ë ¸²"
        Exit Sub
    Else
        vasPrintBuf.PrintOrientation = PrintOrientationPortrait
        vasPrintBuf.Action = 13
    End If
    
    
End Sub

Private Sub cmdWUp_Click()
    Dim lRow As Long
    
    lRow = vasWorkList.ActiveRow
    
    vasWorkList.SwapRange 1, lRow, vasWorkList.MaxCols, lRow, 1, lRow - 1
    SetActiveCell vasWorkList, lRow - 1, 2
'    vasWorkList_Click 2, lRow - 1
End Sub

Private Sub cmdXmlDel_Click()
    Dim FindFile As String

    FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
    If FindFile <> "" Then
        Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     'Àü¼Û¿Ï·á°¡ µÆÀ»¶§ ÆÄÀÏÁö¿ì±â
    End If

End Sub

Private Sub Command1_Click()

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If

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
    '-- ¹ÙÄÚµå
    strBuffer = "D 000201 039903073000126             E01   226H 02    85H 11   9.1  13   141H 18  13.2  21   0.7  24  20.2H "
    
    strBuffer = "DERERBDB"
    strBuffer = "R 003201 0018          1013002058"
    
    strBuffer = "D 003401 0019          1013002058    E      32   1.4  46    26  26  0.81H 01   130  02  3.32L 03  4.29  04   7.3  05   0.5  06   0.1  07   158  09   124H 10   0.7L 11  11.2  12    57  14    39H 15    47H 16    74H 17   259  19   9.1  21   4.7H "
    
    strBuffer = "R 000101 00011013002042"
    
    strBuffer = "D 000101 00011013002042    E012    18  017   129  018    26  "
    
    strBuffer = ""
    strBuffer = strBuffer & "R 000603 0013100825024755"
    
    
    'strBuffer = "D 003401 0019          100825024755  E      32   1.4  46    26  26  0.81H 01   130  02  3.32L 03  4.29  04   7.3  05   0.5  06   0.1  07   158  09   124H 10   0.7L 11  11.2  12    57  14    39H 15    47H 16    74H 17   259  19   9.1  21   4.7H "
'    strBuffer = "D 000103 0100060000034263    0001   7.8  002   4.9  003   2.9  004   1.7  005  1.25  006  0.28  007  0.97  008   117  009    87  010    71  011   250  012   123  013  11.6  014  1.04  015  11.2  016   128  017    84  018    17  019    57  020   257  097    54  "
                 
    strBuffer = "D 001408 006407102000164356    E001   7.5  002   4.7  004    12  005    14  012    80  013   8.1  014  87.2  015   4.8  017   5.1  019   9.4  099   101  "
'
'    strBuffer = "D 001201 00042000186446    E004    22  005    13  008   190  009  76.2  010  35.7  011 139.1  013   0.8  014  98.4  020 124.3  "
'
'    strBuffer = "D 001508 003806272000188274    E004      %?005      %?008      %?009      %?010      %?"
'
    strBuffer = ":n     1   1                              7  1    96   4   185   5   143   6    49  11    18  12    18  14    11 "
    
    strBuffer = ":n    17  17                              8  1   100   4   266   5   232   6    54   9  1.05  11    32  12    35  14    11 "

    Call comEqp_OnComm
        

End Sub

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    Me.Left = 0
    Me.Top = 0
    
    Me.Height = 11520
    Me.Width = 15435
    
    cmdIFClear_Click
    cmdRClear_Click
    lblclear_Click
    
    GetSetup
    
    frmInterface.StatusBar1.Panels(1).Text = gUserID
        
    
    comEqp.CommPort = gSetup.gPort
    comEqp.RTSEnable = gSetup.gRTSEnable
    comEqp.DTREnable = gSetup.gDTREnable
    comEqp.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If comEqp.PortOpen = False Then
        comEqp.PortOpen = True
    End If
    
    If Not Connect_Local Then
        MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
        cn_Local_Flag = False
        Exit Sub
        
    Else
        cn_Local_Flag = True
    End If
    
'    cboGbn.Clear
'    cboGbn.AddItem "ÀüÃ¼"
'    cboGbn.AddItem "°ËÁø"
'    cboGbn.AddItem "Áø·á"
'    cboGbn.ListIndex = 0
    
    '-- osw Ãß°¡
'''    For i = 1 To 1
'''        If Not Connect_PRServer Then
'''            'Cn_Cnt = Cn_Cnt + 1
'''            'If Cn_Cnt = 3 Then
'''            '    If Not Connect_DRServer Then
'''                    MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
'''                    cn_Server_Flag = False
'''                    Exit Sub
'''             '   Else
'''             '       cn_Server_Flag = True
'''             '   End If
'''            'End If
'''        Else
'''            cn_Server_Flag = True
'''        End If
'''    Next
    
    GetExamCode
    
'    SetExamCode
    
    cboPart.AddItem "ÀüÃ¼"
    cboPart.AddItem "¿Ü·¡"
    cboPart.AddItem "°ËÁø"
'    cboPart.AddItem "Ã¤¿ë"
    cboPart.ListIndex = 0

    cboPartR.AddItem "ÀüÃ¼"
    cboPartR.AddItem "¿Ü·¡"
    cboPartR.AddItem "°ËÁø"
    cboPartR.ListIndex = 0

    
    
    dtpToday = Date
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -1), "yyyymmdd")
    
    SQL = "delete from pat_res where transdt < '" & sDate & "'"
    Res = SendQuery(gLocal, SQL)
    
    lblUser.Caption = gUserID
    
    If lblUser.Caption = "" Then
        Call picLogin_Click
    End If
    
'    dtpStartDt.Value = DateAdd("D", -30, Now)
    dtpStartDt.Value = Now 'DateAdd("D", -30, Now)
    dtpStopDt.Value = Now
    
    stInterface.Tab = 0

    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    vasID.MaxRows = 0
'    For i = 1 To 10
'        Call SetText(vasID, CStr(i), i, colBarcode)
'        Call SetText(vasID, "", i, colState)
'    Next
    
'    Winsock1.LocalPort = CInt(5600)
'    Winsock1.Listen
    
End Sub


Private Sub SetExamCode()
    Dim i As Integer
    
    
    With vasRID
        .MaxCols = 7 + UBound(gArrEquip)
        
        For i = 0 To UBound(gArrEquip) - 1
            .Col = colState + (i + 1)
            .Row = -1
            .CellType = CellTypeEdit
            .TypeEditCharSet = TypeEditCharSetASCII
            .TypeEditCharCase = TypeEditCharCaseSetNone
            
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter

            Call SetText(vasRID, gArrEquip(i + 1, 4), 0, 7 + (i + 1))
            .ColWidth(7 + (i + 1)) = 6
            
        Next
        
    End With
    
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

'    Call dce_close_env      ' Server¿Í ¿¬°áÀ» ²÷´Â °÷
    DisConnect_Server
    DisConnect_Local
    Unload Me
    End
End Sub

Private Sub lblDate_DblClick()
    
    If chkSave.Visible = False Then
        chkSave.Visible = True
    Else
        chkSave.Visible = False
    End If
    
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
'   ±â´É : ¿À´õÁ¤º¸ Àü¼Û
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '¼Û½ÅÇÒ µ¥ÀÌÅÍ
    
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
                '## Á¢¼öÁ¤º¸°¡ ¾øÀ»°æ¿ì
                strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "|" & mOrder.Seq & "^" & mOrder.RackNo & _
                            "^" & mOrder.TubePos & "^^SAMPLE^NORMAL|ALL" & _
                            "|R||||||C||||||||||||||Q" & vbCr & ETX
                intSndPhase = 5
            
            Else
                If mOrder.IsSending = False Then   '## ÃÖÃÊ º¸³¾¶§
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
                Else                        '## ³²Àº ¹®ÀÚ¿­ÀÌ ÀÖÀ»¶§
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
'   ±â´É : ÇØ´ç ¹®ÀÚ¿­ÀÇ CheckSumÀ» ±¸ÇÔ
'   ÀÎ¼ö :
'       - pMsg : ¹®ÀÚ¿­
'   ¹ÝÈ¯ : CheckSum
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

'-- Áö±Ý³¯Â¥¿Í °Ë»çÀÏÀÚ ºñ±³ÇÑ´Ù
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
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Static ChkSumCnt As Long
    
    Select Case comEqp.CommEvent
        Case comEvReceive
            
            Buffer = comEqp.Input
            'Buffer = strBuffer
            
            SetRawData "[Rx]" & Buffer
            lngBufLen = Len(Buffer)
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case BufChar
                    Case STX
                        strBuffer = ""
                        
                    Case ETX
                         Call EditRcvData
                         strBuffer = ""
                    
                    Case Else
                        strBuffer = strBuffer & BufChar
                End Select
            Next i

        Case comEvSend
        Case comEvCTS
            EVMsg$ = "CTS º¯°æ °¨Áö"
        Case comEvDSR
            EVMsg$ = "DSR º¯°æ °¨Áö"
        Case comEvCD
            EVMsg$ = "CD º¯°æ °¨Áö"
        Case comEvRing
            EVMsg$ = "ÀüÈ­ º§ÀÌ ¿ï¸®´Â Áß"
        Case comEvEOF
            EVMsg$ = "EOF °¨Áö"

        '¿À·ù ¸Þ½ÃÁö
        Case comBreak
            ERMsg$ = "Áß´Ü ½ÅÈ£ ¼ö½Å"
        Case comCDTO
            ERMsg$ = "¹Ý¼ÛÆÄ °ËÃâ ½Ã°£ ÃÊ°ú"
        Case comCTSTO
            ERMsg$ = "CTS ½Ã°£ ÃÊ°ú"
        Case comDCB
            ERMsg$ = "DCB °Ë»ö ¿À·ù"
        Case comDSRTO
            ERMsg$ = "DSR ½Ã°£ ÃÊ°ú"
        Case comFrame
            ERMsg$ = "ÇÁ·¹ÀÌ¹Ö ¿À·ù"
        Case comOverrun
            ERMsg$ = "ÆÐ¸®Æ¼ ¿À·ù"
        Case comRxOver
            ERMsg$ = "¼ö½Å ¹öÆÛ ÃÊ°ú"
        Case comRxParity
            ERMsg$ = "ÆÐ¸®Æ¼ ¿À·ù"
        Case comTxFull
            ERMsg$ = "Àü¼Û ¹öÆÛ¿¡ ¿©À¯°¡ ¾øÀ½"
        Case Else
            ERMsg$ = "¾Ë ¼ö ¾ø´Â ¿À·ù ¶Ç´Â ÀÌº¥Æ®"
    End Select


End Sub


Public Sub sndMore()
    Dim strSndMsg As String
    
    strSndMsg = ">"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & vbCrLf
    
End Sub

Public Sub sndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & vbCrLf
    
End Sub

'-----------------------------------------------------------------------------'
'   ±â´É : ÇØ´ç ¹ÙÄÚµå¹øÈ£¿¡ ´ëÇÑ Á¢¼öÁ¤º¸ Á¶È¸, tblReady, tblResult¿¡ Ç¥½Ã
'   ÀÎ¼ö :
'       - pBarNo : ¹ÙÄÚµå¹øÈ£
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    
    intRow = -1
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colPID)) = pBarNo Then
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
    
    Call SetText(vasID, "Order", intRow, colState)
    Call SetText(vasID, pBarNo, intRow, colBarNo)
    Call vasActiveCell(vasID, intRow, colBarNo)
    Call ClearSpread(vasRes)
    
    Call GetSampleInfoW(intRow)                            '5,6,7,8
    
    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)
    
    '-- ±âÁ¸ °Ë»çÇß´ø ¹ÙÄÚµå°¡ ´Ù½Ã ¿Ã¶ó¿Ã °æ¿ì À§Ä¡¸¦ ¸øÃ£´Â´Ù.
    '-- intRow Ãß°¡
'    strItems = GetGetEquipExamCode_AU480(gEquip, pBarNo, intRow)
'
'    If Trim(strItems) = "" Then
'        mOrder.NoOrder = True
'        mOrder.Order = ""
'        'S 003401 0019          1013001918    E
'        'comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
'        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & ETX
'        Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & ETX
'
'        SetRawData "[Tx]" & STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & ETX
'
'    Else
'        mOrder.NoOrder = False
'        mOrder.Order = strItems
'        'S 003401 0019          1013001918    E      01020304050607091011121415161719212632
'        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
'        'comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E012" & ETX
'
'
'        Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
'        SetRawData "[Tx]" & STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
'
'    End If

End Sub

'-----------------------------------------------------------------------------'
'   ±â´É :
'   ÀÎ¼ö :
'       - pBarNo : ¹ÙÄÚµå¹øÈ£
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    
    intRow = -1
    For i = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, i, colPID)) = pBarNo Then
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
    
    If Mid(pBarNo, 1, 2) = "QC" Then
        Call SetText(vasID, Format(Now, "yyyy-mm-dd"), intRow, 2)
        Call SetText(vasID, pBarNo, intRow, colPID)
        Call SetText(vasID, "QC", intRow, colPName)
    Else
        'Call SetText(vasID, pBarNo, intRow, colBarNo)             '2 Barcode
    End If
    
    Call vasActiveCell(vasID, intRow, colPID)
    
    Call ClearSpread(vasRes)
    
    Call GetSampleInfoW(intRow)                                '5,6,7,8
    
    gRow = intRow
    
    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)

End Sub

'-----------------------------------------------------------------------------'
'   ±â´É : Àåºñ·ÎºÎ ¼ö½ÅÇÑ µ¥ÀÌÅÍ ÆíÁý
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim strRcvBuf    As String   '¼ö½ÅÇÑ Data
    Dim strType      As String   '¼ö½ÅÇÑ Record Type
    Dim strBarNo     As String   '¼ö½ÅÇÑ ¹ÙÄÚµå¹øÈ£
    Dim strSeq       As String   '¼ö½ÅÇÑ Sequence
    Dim strRackNo    As String   '¼ö½ÅÇÑ Rack Or Disk No
    Dim strTubePos   As String   '¼ö½ÅÇÑ Tube Position
    Dim strIntBase   As String   '¼ö½ÅÇÑ Àåºñ±âÁØ °Ë»ç¸í
    Dim strResult    As String   '¼ö½ÅÇÑ °á°ú
    Dim strQCResult  As String   '¼ö½ÅÇÑ °á°ú(QC)
    Dim strFlag      As String   '¼ö½ÅÇÑ Abnormal Flag
    Dim strComm      As String   '¼ö½ÅÇÑ Comment
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
    Dim strSndBuffer As String
    Dim intCol As Integer
    
    Dim strFunc       As String
    Dim strFunction   As String
    Dim strDate       As String
    Dim strSendData   As String
    Dim strSndMsg     As String
    Dim strExamCode() As String
    
    Dim i             As Integer
    
    strRcvBuf = strBuffer
    
    strType = Mid$(strRcvBuf, 1, 1)
    If IsNumeric(strType) Then
        strType = Mid$(strRcvBuf, 2, 1)
    End If
    
    Select Case strType
        Case ">", "?", "@"      'ANY ¼ö½Å
            Sleep (100)
            Call sndMore        'MOR Send
            Do
            '   DoEvents
            Loop Until comEqp.OutBufferCount = 0
        
        Case "?", "@"           'REP ¼ö½Å
            Sleep (100)
            Call sndMore        'MOR Send
            Do
            '   DoEvents
            Loop Until comEqp.OutBufferCount = 0
        
        Case ">", "?", "@"      'SUS ¼ö½Å
            Sleep (100)
            Call sndMore        'MOR Send
            Do
            '   DoEvents
            Loop Until comEqp.OutBufferCount = 0
        
        Case ";"                'SPE  Àü¼Û(¿À´õ¿äÃ»)
            ';N     1   1                            
            'strFunction = Mid(strRcvBuf, 2, 40)
            
            ';N     1   1 213296                      37111111100000000000000000000000000000000000
';N     1   1     
            strFunction = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
            
            'sFunc = Replace(sFunc, String(13, "#"), Left(pSampleInfo.ID & Space(13), 13))

            strFunc = Mid(strRcvBuf, 2, 1)              'N
            
           If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                Exit Sub
            End If
            
            strSeq = Mid(strRcvBuf, 4, 5)               '    1
            strRackNo = Mid(strRcvBuf, 9, 1)            '
            strTubePos = Mid(strRcvBuf, 10, 3)          '  1
            'strBarNo = Trim(Mid(strRcvBuf, 13, 13))
            strBarNo = Trim(Mid(strRcvBuf, 14, 13))
            
            
            '-- ¿öÅ©Á¶È¸¸¦ ¾ÈÇßÀ¸¸é
            If vasID.DataRowCnt = 0 Then
                Exit Sub
            End If
            
            '-- Æ÷Áö¼Ç ¹øÈ£·Î Ã­Æ® ¹øÈ£ Ã£±â
            For i = 1 To vasID.DataRowCnt
                If Trim(GetText(vasID, i, colOCnt)) = Val(strTubePos) Then
                    strBarNo = Trim(GetText(vasID, i, colPID))
                    'strSeq = Trim(GetText(vasID, i, colPID))
                    gRow = i
                    Exit For
                End If
            Next
            
            If strBarNo = "" Or gRow < 0 Then Exit Sub
            
            With mOrder
                .BarNo = Trim(strBarNo)
                .Seq = Trim(strSeq)
            End With
            
            Call GetOrder(strBarNo)
            
            strFunction = Replace(strFunction, String(13, "#"), Left(mOrder.BarNo & Space(13), 13))
            
            strSendData = String$(88, "0")
            
            '-- °¡Á®¿Â °Ë»çÄÚµåÀÇ Ã¤³Î Ã£±â
            SQL = ""
            SQL = SQL & "SELECT Distinct EQUIPCODE "
            SQL = SQL & "  FROM equipexam "
            SQL = SQL & " WHERE EQUIPNO  = '" & Trim(gEquip) & "' "
            SQL = SQL & "   AND EXAMCODE in (" & Trim(gOrderExam) & ")"
            
            Res = GetDBSelectRow(gLocal, SQL)
            
            For i = 0 To UBound(gReadBuf)
                If gReadBuf(i) <> "" Then
                    ReDim Preserve strExamCode(ii)
                    strExamCode(ii) = Trim(gReadBuf(i))
                    ii = ii + 1
                Else
                    Exit For
                End If
            Next
            
            For i = 0 To UBound(strExamCode)
                If strExamCode(i) <> "" Then
                    If strExamCode(i) <> "99" Then
                        Mid(strSendData, strExamCode(i), 1) = "1"
                    End If
                End If
            Next
            
            strSndMsg = ";" & strFunction
            strSndMsg = strSndMsg & " 37"
            strSndMsg = strSndMsg & Mid(strSendData, 1, 37)
            strSndMsg = strSndMsg & "00000"
            
            'COMMENT¶õ¿¡ BARCODE Ç¥½Ã
            'strSndMsg = strSndMsg & "100000" & Left(strBarNo & Space(30), 30)
            
            
            Sleep (100)
            '-- SPE Send(¿À´õÀü¼Û)
            'comEqp.Output = STX & strSndMsg & ETX & GetChkSum(strSndMsg) & vbCr
            comEqp.Output = STX & strSndMsg & ETX & vbCrLf
            
            SetRawData "[Tx]" & STX & strSndMsg & ETX & vbCrLf
            
            
        ' FR1 to FR9 (°Ë»çÇ×¸ñ 25°³ ÀÌ»óÀÏ °æ¿ì Ã³¸®)
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
            strFunc = Mid(strRcvBuf, 2, 1)
            
            If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                Sleep (100)
                Call sndMore        'MOR Send
                Do
                '   DoEvents
                Loop Until comEqp.OutBufferCount = 0
                Exit Sub
            End If
            
            If strFunc = "K" Or strFunc = "L" Or strFunc = "F" Then
                Call sndMore        'MOR Send
                Exit Sub
            End If
            
            If strFunc <> "@" And strFunc <> "M" Then
                strRackNo = Mid(strRcvBuf, 9, 1)
                strTubePos = Trim(Mid(strRcvBuf, 10, 3))
                'strBarNo = Trim(Mid(strRcvBuf, 13, 13))
                strBarNo = Trim(Mid(strRcvBuf, 14, 13))
                gRow = 0
                '-- Æ÷Áö¼Ç ¹øÈ£·Î ¹ÙÄÚµå ¹øÈ£ Ã£±â
                For i = 1 To vasID.DataRowCnt
                    If Trim(GetText(vasID, i, colOCnt)) = Val(strTubePos) Then
                        strBarNo = Trim(GetText(vasID, i, colPID))
                        gRow = i
                        Exit For
                    End If
                Next
                
                If strBarNo = "" Or gRow <= 0 Then Exit Sub
                
                With mResult
                    .BarNo = Trim(strBarNo)
                    .Seq = Trim(strSeq)
                    .RackNo = Trim(strRackNo)
                    .TubePos = Trim(strTubePos)
                    '.RsltDate = Format(Now, "yyyymmddhhmmss")
                    '.RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                End With
            
                Call SetPatInfo(strBarNo)

        
                For i = 44 To Len(strRcvBuf) Step 10
                'For i = 52 To Len(strRcvBuf) Step 10
                    strIntBase = Trim(Mid(strRcvBuf, i, 3))
'                    strIntBase = Format(strIntBase, "00")
                    strResult = Trim(Mid(strRcvBuf, i + 3, 6))
                    
                    '-- CRP
'                    If strIntBase = "13" Then
'                        If IsNumeric(strResult) Then
'                            If CCur(strResult) < 0.1 Then
'                                strResult = "<0.1"
'                            End If
'                        End If
'                    End If
                    
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO,REFLOW,REFHIGH "
                        SQL = SQL & "  FROM EQUIPEXAM"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        '-- ¿À´õ ÀÖÀ» °æ¿ì
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '¼Ò¼öÁ¡ Ã³¸®, °á°ú ÇüÅÂ Ã³¸®
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText vasID, "1", gRow, colCheckBox                  'Ã¼Å©
                            'SetText vasID, strResult, gRow, colA1c                  '°á°ú
                            'SetText vasID, strComm, gRow, colA1c + 1                'Flag
                            SetText vasID, "Result", gRow, colState                 'ÁøÇà»óÅÂ
                            '-- °á°ú List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                            SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                            
                            
                            SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                            
                            vasRes.Row = lsResRow
                            vasRes.Col = colResult
                            vasRes.FontBold = False
                            vasRes.ForeColor = vbBlack
                            
                            
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                            'SetText vasRes, strComm, lsResRow, 7                    'Flag
                            
                            '-- ·ÎÄÃ ÀúÀå
                            SetLocalDB_New gRow, lsResRow, "1", lsEquipRes
                            
                            lsResult_Buff = ""
                            strState = "R"
                            
                        '-- ¿À´õ ¾øÀ» °æ¿ì
                        Else
                        
                                  SQL = "Select examcode, examname, seqno ,REFLOW,REFHIGH "
                            SQL = SQL & "  From equipexam"
                            SQL = SQL & " Where equipno = '" & gEquip & "' "
                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                            SQL = SQL & "   and examtype = 'C' "
                            Res = GetDBSelectColumn(gLocal, SQL)
                            
                            If Res > 0 Then
                                lsExamCode = Trim(gReadBuf(0))
                                lsExamName = Trim(gReadBuf(1))
                                lsSeqNo = Trim(gReadBuf(2))
                                
                                lsResRow = vasRes.DataRowCnt + 1
                                If vasRes.MaxRows < lsResRow Then
                                    vasRes.MaxRows = lsResRow
                                End If
                                
                                '¼Ò¼öÁ¡ Ã³¸®, °á°ú ÇüÅÂ Ã³¸®
                                lsEquipRes = strResult
                                strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                '-- Work List
                                'SetText vasID, "0", gRow, colCheckBox                  'Ã¼Å©
                                'SetText vasID, strResult, gRow, colA1c                  '°á°ú
                                'SetText vasID, strComm, gRow, colA1c + 1                'Flag
                                SetText vasID, "Result", gRow, colState                 'ÁøÇà»óÅÂ
                                '-- °á°ú List
                                SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                                SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                                                                
                                SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                                
                                vasRes.Row = lsResRow
                                vasRes.Col = colResult
                                vasRes.FontBold = False
                                vasRes.ForeColor = vbBlack
        
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                '-- ·ÎÄÃ ÀúÀå
                                SetLocalDB_New gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                strState = "R"
                            End If
                        End If
                    End If
                Next
                
                If MnTransAuto.Checked = True And strState = "R" Then
                    
                    Res = SaveTransDataW(gRow)
                    
                    If Res = -1 Then
                        '-- ÀúÀå ½ÇÆÐ
                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                        SetText vasID, "Failed", gRow, colState
                    Else
                        '-- ÀúÀå ¼º°ø
                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                        SetText vasID, "Trans", gRow, colState
                        
                        SQL = " Update pat_res Set " & vbCrLf & _
                              "  transdt = '" & Format(Now, "yyyymmdd") & "', " & vbCrLf & _
                              "  transyn = '2' " & vbCrLf & _
                              "  Where barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' " & _
                              "    and examtype = 'C'"
                              
                        Res = SendQuery(gLocal, SQL)
                        If Res = -1 Then
                            SaveQuery SQL
                            Exit Sub
                        End If
                        SetText vasID, "0", gRow, colCheckBox
            
                    End If
                    strState = ""
                End If
                
                Call sndMore
            End If
            
        ' °á°ú END
        Case ":"
            strFunc = Mid(strRcvBuf, 2, 1)
            
            If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                Sleep (100)
                Call sndMore        'MOR Send
                Do
                '   DoEvents
                Loop Until comEqp.OutBufferCount = 0
                Exit Sub
            End If
            
            If strFunc = "K" Or strFunc = "L" Or strFunc = "F" Then
                Call sndMore        'MOR Send
                Exit Sub
            End If
            
            
            
            Call sndMore            'MOR Send
            
            
            If strFunc <> "@" And strFunc <> "M" Then
                strRackNo = Mid(strRcvBuf, 9, 1)
                strTubePos = Trim(Mid(strRcvBuf, 10, 3))
                'strBarNo = Trim(Mid(strRcvBuf, 13, 13))
                strBarNo = Trim(Mid(strRcvBuf, 14, 13))
                gRow = 0
                'QC
                If strFunc = "F" Then
                    vasID.MaxRows = vasID.MaxRows + 1
                    strBarNo = Trim(Mid(strRcvBuf, 6, 10))
                    strBarNo = "QC" & strBarNo
                    gRow = vasID.MaxRows
                Else
                    '-- Æ÷Áö¼Ç ¹øÈ£·Î ¹ÙÄÚµå ¹øÈ£ Ã£±â
                    For i = 1 To vasID.DataRowCnt
    '                    vasID.Row = i
    '                    vasID.Col = 0
    '                    vasID.Text = ""
    '                    vasID.colhea
                        If Trim(GetText(vasID, i, colOCnt)) = Val(strTubePos) Then
                            strBarNo = Trim(GetText(vasID, i, colPID))
                            gRow = i
                            Exit For
                        End If
                    Next
                End If
                
                If strBarNo = "" Or gRow <= 0 Then Exit Sub
                
                With mResult
                    .BarNo = Trim(strBarNo)
                    .Seq = Trim(strSeq)
                    .RackNo = Trim(strRackNo)
                    .TubePos = Trim(strTubePos)
                    '.RsltDate = Format(Now, "yyyymmddhhmmss")
                    '.RsltSeq = getMaxTestNum(Format(dtpToday, "yyyymmdd"))
                End With
            
                Call SetPatInfo(strBarNo)

        
                For i = 44 To Len(strRcvBuf) Step 10
                'For i = 51 To Len(strRcvBuf) Step 10
                    strIntBase = Trim(Mid(strRcvBuf, i, 3))
                    'strIntBase = Format(strIntBase, "00")
                    strResult = Trim(Mid(strRcvBuf, i + 3, 6))
                    
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO,REFLOW,REFHIGH "
                        SQL = SQL & "  FROM EQUIPEXAM"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                        
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        '-- ¿À´õ ÀÖÀ» °æ¿ì
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '¼Ò¼öÁ¡ Ã³¸®, °á°ú ÇüÅÂ Ã³¸®
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText vasID, "1", gRow, colCheckBox                  'Ã¼Å©
                            'SetText vasID, strResult, gRow, colA1c                  '°á°ú
                            'SetText vasID, strComm, gRow, colA1c + 1                'Flag
                            SetText vasID, "Result", gRow, colState                 'ÁøÇà»óÅÂ
                            '-- °á°ú List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                            SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                            
                            
                            SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                            
                            vasRes.Row = lsResRow
                            vasRes.Col = colResult
                            vasRes.FontBold = False
                            vasRes.ForeColor = vbBlack
                            
                            
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                            'SetText vasRes, strComm, lsResRow, 7                    'Flag
                            
                            '-- ·ÎÄÃ ÀúÀå
                            SetLocalDB_New gRow, lsResRow, "1", lsEquipRes
                            
                            lsResult_Buff = ""
                            strState = "R"
                            
                        '-- ¿À´õ ¾øÀ» °æ¿ì
                        Else
                        
                                  SQL = "Select examcode, examname, seqno ,REFLOW,REFHIGH "
                            SQL = SQL & "  From equipexam"
                            SQL = SQL & " Where equipno = '" & gEquip & "' "
                            SQL = SQL & "   and equipcode = '" & strIntBase & "' "
                            SQL = SQL & "   and examtype = 'C' "
                            Res = GetDBSelectColumn(gLocal, SQL)
                            
                            If Res > 0 Then
                                lsExamCode = Trim(gReadBuf(0))
                                lsExamName = Trim(gReadBuf(1))
                                lsSeqNo = Trim(gReadBuf(2))
                                
                                lsResRow = vasRes.DataRowCnt + 1
                                If vasRes.MaxRows < lsResRow Then
                                    vasRes.MaxRows = lsResRow
                                End If
                                
                                '¼Ò¼öÁ¡ Ã³¸®, °á°ú ÇüÅÂ Ã³¸®
                                lsEquipRes = strResult
                                strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                '-- Work List
                                'SetText vasID, "0", gRow, colCheckBox                  'Ã¼Å©
                                'SetText vasID, strResult, gRow, colA1c                  '°á°ú
                                'SetText vasID, strComm, gRow, colA1c + 1                'Flag
                                SetText vasID, "Result", gRow, colState                 'ÁøÇà»óÅÂ
                                '-- °á°ú List
                                SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                                SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                                                                
                                SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                                
                                vasRes.Row = lsResRow
                                vasRes.Col = colResult
                                vasRes.FontBold = False
                                vasRes.ForeColor = vbBlack
        
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                
                                If Mid(strBarNo, 1, 2) = "QC" Then
                                    '-- ·ÎÄÃ ÀúÀå
                                    SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                Else
                                    '-- ·ÎÄÃ ÀúÀå
                                    SetLocalDB_New gRow, lsResRow, "1", lsEquipRes
                                End If
                                
                                lsResult_Buff = ""
                                strState = "R"
                            End If
                        End If
                    End If
                Next
                
                vasRes.RowHeight(-1) = 12
                
                If MnTransAuto.Checked = True And strState = "R" Then
                    
                    Res = SaveTransDataW(gRow)
                    
                    If Res = -1 Then
                        '-- ÀúÀå ½ÇÆÐ
                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                        SetText vasID, "Failed", gRow, colState
                    Else
                        '-- ÀúÀå ¼º°ø
                        SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                        SetText vasID, "Trans", gRow, colState
                        
                        SQL = " Update pat_res Set " & vbCrLf & _
                              "  transdt = '" & Format(Now, "yyyymmdd") & "', " & vbCrLf & _
                              "  transyn = '2' " & vbCrLf & _
                              "  Where barcode = '" & Trim(GetText(vasID, gRow, colBarcode)) & "' " & _
                              "    and examtype = 'C'"
                              
                        Res = SendQuery(gLocal, SQL)
                        If Res = -1 Then
                            SaveQuery SQL
                            Exit Sub
                        End If
                        SetText vasID, "0", gRow, colCheckBox
            
                    End If
                    strState = ""
                End If
            End If
            
        Case Else
                Call sndMore
        
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
        If IsNumeric(sEquipRes) Then
            If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
                sResFlag = ""
            ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
                sResFlag = "H"
            ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
                sResFlag = "L"
            End If
        End If
    End If
    
    gsFlag = sResFlag
    SetResult = sResult
    
End Function

' asRow1 = Work List
' asRow2 = °á°ú List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = Format(dtpToday, "yyyymmdd")


'    SQL = ""
'    SQL = SQL & "UPDATE PAT_RES SET "
'    SQL = SQL & " RESULT = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf
'    SQL = SQL & " RSLTDATE = '" & Format(CDate(dtpToday.Value), "yyyymmdd") & "' " & vbCrLf
'    SQL = SQL & " WHERE BARCODE = '" & Trim(GetText(vasID, asRow1, 4)) & "' " & vbCrLf
'    SQL = SQL & "   AND COMEXAMID = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf
'    SQL = SQL & "   AND EXAMID = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'" & vbCrLf
'    SQL = SQL & "   AND EXAMTYPE = 'C'"

    SQL = ""
    SQL = "DELETE FROM PAT_RES " & vbCrLf
    SQL = SQL & "WHERE RSLTDATE = '" & Format(Now, "YYYYMMDD") & "' " & vbCrLf
    SQL = SQL & "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colPID)) & "' " & vbCrLf
    SQL = SQL & "  AND COMEXAMID = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf
    SQL = SQL & "  AND EXAMID = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
'    SQL = ""
'    SQL = SQL & "INSERT INTO PAT_RES("
'    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
'                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
'                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID,EXAMTYPE) " & vbCrLf
'    SQL = SQL & "VALUES("
'    SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
'    SQL = SQL & "'" & gEquip & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colDISK)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPName)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colSex)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colAge)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', "
'    SQL = SQL & "'0', "
'    SQL = SQL & "'C',"
'    SQL = SQL & "'" & gIFUser & "')"
'
    
          SQL = " insert into pat_res("
    SQL = SQL & "Company,HospCode,ChartNo, "
    SQL = SQL & "PatName,PatSex,PatAge,PatJumin,PatNo,"
    SQL = SQL & "CommDate,ExamNo,ExamID,ComExamID, "
    SQL = SQL & "Specimen,Result,Reference,Remark,RsltDate,IOFlag,BarCode,ExamType)"
    SQL = SQL & " values ("
    SQL = SQL & "'ACK',"
    SQL = SQL & "'41343051',"
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "',"
    SQL = SQL & "'QC',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'" & Format(Now, "yyyymmdd") & "',"
    SQL = SQL & "'',"
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "',"
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "',"
    SQL = SQL & "'',"
    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'" & Format(Now, "yyyymmdd") & "',"
    SQL = SQL & "'',"
    SQL = SQL & "'',"
    SQL = SQL & "'C')"
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Function SetLocalDB_New(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = Format(dtpToday, "yyyymmdd")

'    SQL = ""
'    SQL = "DELETE FROM PAT_RES " & vbCrLf & _
'          "WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf & _
'          "  AND EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'          "  AND BARCODE = '" & Trim(GetText(vasID, asRow1, colBarcode)) & "' " & vbCrLf & _
'          "  AND EQUIPCODE = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
'          "  AND EXAMCODE = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
'
'    Res = SendQuery(gLocal, SQL)
'
'    If Res = -1 Then
'        SaveQuery SQL
'        Exit Function
'    End If
    
'    SQL = ""
'    SQL = SQL & "INSERT INTO PAT_RES("
'    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
'                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
'                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
'    SQL = SQL & "VALUES("
'    SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
'    SQL = SQL & "'" & gEquip & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colDISK)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPName)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colSex)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colAge)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamCode)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colSeq)) & "', " & vbCrLf
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', "
'    SQL = SQL & "'" & Trim(GetText(vasRes, asRow2, colExamName)) & "', "
'    SQL = SQL & "'0', "
'    SQL = SQL & "'" & gIFUser & "')"
    
    SQL = ""
    SQL = SQL & "UPDATE PAT_RES SET "
    SQL = SQL & " RESULT = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf
    SQL = SQL & " RSLTDATE = '" & Format(CDate(dtpToday.Value), "yyyymmdd") & "' " & vbCrLf
    SQL = SQL & " WHERE CHARTNO = '" & Trim(GetText(vasID, asRow1, 5)) & "' " & vbCrLf
    SQL = SQL & "   AND COMEXAMID = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf
    SQL = SQL & "   AND EXAMID = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'" & vbCrLf
    SQL = SQL & "   AND EXAMTYPE = 'C'"
    
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
    sMsg = "°Ë»çÀÚ¸¦ ÀÔ·ÂÇØÁÖ¼¼¿ä."
    lblUser.Caption = InputBox(sMsg, "°Ë»çÀÚ ÀÔ·Â")

End Sub



Private Sub txtNum_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer

    If KeyAscii = 13 Then
        With vasID
            For intRow = .ActiveRow To .DataRowCnt
                SetText vasID, txtNum, intRow, colOCnt 'colSpecNo

                txtNum = Val(txtNum) + 1
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
    Dim strChart As String
    Dim strDate As String
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    strChart = Trim(GetText(vasID, Row, 5))
    lsID = Trim(GetText(vasID, Row, 4))
    strDate = Trim(GetText(vasID, Row, 2))
    strDate = Replace(strDate, "-", "")
    
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasID, Row, colPID))
    lblBarcode(0).Caption = strChart 'lsID
    lblPname(0).Caption = Trim(GetText(vasID, Row, colPName))
    'Local¿¡¼­ ºÒ·¯¿À±â
    ClearSpread vasRes
    
    'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
          SQL = "SELECT a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.EXAMNO, a.TRANSYN, b.reflow,b.refhigh, b.seqno " & vbCrLf
    SQL = SQL & "  FROM PAT_RES a, EQUIPEXAM b " & vbCrLf
    SQL = SQL & " WHERE a.EXAMID = b.EXAMCODE " & vbCrLf
    SQL = SQL & "   AND a.COMEXAMID = b.EQUIPCODE " & vbCrLf
    'SQL = SQL & "   AND a.EQUIPNO = '" & gEquip & "' " & vbCrLf
'    SQL = SQL & "   AND a.BARCODE = '" & lsID & "' " & vbCrLf
    SQL = SQL & "   AND a.COMMDATE = '" & strDate & "' " & vbCrLf
    SQL = SQL & "   AND a.CHARTNO = '" & strChart & "' " & vbCrLf
    SQL = SQL & "   AND a.EXAMTYPE = 'C' " & vbCrLf
    SQL = SQL & " GROUP BY a.EXAMNO, a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.TRANSYN, b.reflow,b.refhigh, b.seqno "
    SQL = SQL & " ORDER BY b.seqno * 10"
    Res = GetDBSelectVas_Ref(gLocal, SQL, vasRes)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    Else
        vasRes.RowHeight(-1) = 12
    End If

    vasRes.MaxRows = vasRes.DataRowCnt

End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    Dim lsNM As String
    Dim lsPid As String
    Dim i As Integer

    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If

        lsID = Trim(GetText(vasID, iRow, colBarcode))
        lsPid = Trim(GetText(vasID, iRow, colPID))
        lsNM = Trim(GetText(vasID, iRow, colPName))
        
        If MsgBox(lsNM & "À» Áö¿ì½Ã°Ú½À´Ï±î?", vbInformation + vbYesNo, "¾Ë¸²") = vbNo Then
            Exit Sub
        End If
'
'        SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " AND BARCODE = '" & lsID & "' " & vbCrLf & _
'              " AND PID = '" & lsPid & "' " & vbCrLf & _
'              " AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' " & vbCrLf & _
'              " AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' " & vbCrLf & _
'              " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'        Res = SendQuery(gLocal, SQL)
'
'        If Res = -1 Then
'            SaveQuery SQL
'            Exit Sub
'        End If

        DeleteRow vasID, iRow, iRow
        vasRes.MaxRows = 0

        vasID.MaxRows = vasID.MaxRows - 1

'    ElseIf KeyCode = 13 Then
'
'        GetSampleInfoW (iRow)
'
'        lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'        'Local¿¡¼­ ºÒ·¯¿À±â
'        ClearSpread vasTemp
'
'        'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
'        SQL = "SELECT EQUIPCODE, EXAMCODE, EXAMNAME, SEQNO " & vbCrLf & _
'              "  FROM EQUIPEXAM " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " ORDER BY SEQNO "
'
'        Res = GetDBSelectVas(gLocal, SQL, vasTemp)
'        If Res = -1 Then
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
'                Res = SendQuery(gLocal, SQL)
'            Next i
'
'            SQL = " DELETE FROM PAT_RES " & vbCrLf & _
'                  " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'                  " AND BARCODE = '" & lblChangeBar.Caption & "' " & vbCrLf & _
'                  " AND PID = '" & lblChangePID.Caption & "' " & vbCrLf & _
'                  " AND DISKNO = '" & Trim(GetText(vasID, iRow, colRack)) & "' " & vbCrLf & _
'                  " AND POSNO = '" & Trim(GetText(vasID, iRow, colPos)) & "' " & vbCrLf & _
'                  " AND EXAMDATE = '" & Format(dtpToday.Value, "yyyymmdd") & "' "
'            Res = SendQuery(gLocal, SQL)
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
'                Res = SendQuery(gLocal, SQL)
'            Next i
'        End If
'        SetText vasID, "Result", gRow, colState

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
    Dim lsBarcode As String
    Dim i As Integer
    
    If Row < 1 Or Row > vasRID.DataRowCnt Then
        Exit Sub
    End If
    
    lsBarcode = Trim(GetText(vasRID, Row, 3))
    lsID = Trim(GetText(vasRID, Row, 4))
    lblChangeBar.Caption = lsID
    lblBarcode(1).Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasRID, Row, colPID))
    lblPname(1).Caption = Trim(GetText(vasRID, Row, colPName))
    lblRrow.Caption = Row
    'Local¿¡¼­ ºÒ·¯¿À±â
    ClearSpread vasRRes
    
    'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
          SQL = "SELECT a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.EXAMNO, a.TRANSYN, b.reflow,b.refhigh " & vbCrLf
    SQL = SQL & "  FROM PAT_RES a, EQUIPEXAM b " & vbCrLf
    SQL = SQL & " WHERE a.EXAMID = b.EXAMCODE " & vbCrLf
    SQL = SQL & "   AND a.COMEXAMID = b.EQUIPCODE " & vbCrLf
    SQL = SQL & "   AND a.BARCODE = '" & lsBarcode & "' " & vbCrLf
    SQL = SQL & "   AND a.CHARTNO = '" & lsID & "' " & vbCrLf
    SQL = SQL & "   AND a.EXAMTYPE = 'C' " & vbCrLf
    SQL = SQL & " GROUP BY a.EXAMNO, a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.TRANSYN, b.reflow,b.refhigh "
    
    Res = GetDBSelectVas_Ref(gLocal, SQL, vasRRes)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    Else
        vasRRes.RowHeight(-1) = 12
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
'        'Local¿¡¼­ ºÒ·¯¿À±â
'        ClearSpread vasTemp
'
'        'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
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
'        If MsgBox("ÇØ´ç È¯ÀÚ°á°ú¸¦ »èÁ¦ÇÏ½Ã°Ú½À´Ï±î?", vbInformation + vbYesNo, "¾Ë¸²") = vbNo Then
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



Private Sub vasWorkList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim lsID As String
    Dim strChart As String
    Dim strDate As String
    
    If Row < 1 Or Row > vasWorkList.DataRowCnt Then
        Exit Sub
    End If
    
    strChart = Trim(GetText(vasWorkList, Row, 5))
    lsID = Trim(GetText(vasWorkList, Row, 4))
    strDate = Trim(GetText(vasWorkList, Row, 2))
    strDate = Replace(strDate, "-", "")
    
    lblChangeBar.Caption = lsID
    lblChangePID.Caption = Trim(GetText(vasWorkList, Row, colPID))
    lblBarcode(0).Caption = strChart 'lsID
    lblPname(0).Caption = Trim(GetText(vasWorkList, Row, colPName))
    'Local¿¡¼­ ºÒ·¯¿À±â
    ClearSpread vasRes
    
    'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
          SQL = "SELECT a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.EXAMNO, a.TRANSYN, b.reflow,b.refhigh, b.seqno " & vbCrLf
    SQL = SQL & "  FROM PAT_RES a, EQUIPEXAM b " & vbCrLf
    SQL = SQL & " WHERE a.EXAMID = b.EXAMCODE " & vbCrLf
    SQL = SQL & "   AND a.COMEXAMID = b.EQUIPCODE " & vbCrLf
    'SQL = SQL & "   AND a.EQUIPNO = '" & gEquip & "' " & vbCrLf
'    SQL = SQL & "   AND a.BARCODE = '" & lsID & "' " & vbCrLf
    SQL = SQL & "   AND a.COMMDATE = '" & strDate & "' " & vbCrLf
    SQL = SQL & "   AND a.CHARTNO = '" & strChart & "' " & vbCrLf
    SQL = SQL & "   AND a.EXAMTYPE = 'C' " & vbCrLf
    SQL = SQL & " GROUP BY a.EXAMNO, a.COMEXAMID, a.EXAMID, b.EXAMNAME, a.RESULT, a.TRANSYN, b.reflow,b.refhigh, b.seqno "
    SQL = SQL & " ORDER BY b.seqno * 10"
    Res = GetDBSelectVas_Ref(gLocal, SQL, vasRes)
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    Else
        vasRes.RowHeight(-1) = 12
    End If

    vasRes.MaxRows = vasRes.DataRowCnt
End Sub

Private Sub vasWorkList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim pGrid_Point As Integer
    Dim sBarcode As String
    Dim sChartNo As String
    
    If Row = 0 Then Exit Sub
    
    With vasWorkList
        .Col = Col
        .Row = Row
        .Col = 5
        pGrid_Point = SeqSearch(vasID, Trim(.Text), 5)

        If pGrid_Point = 0 Then
            pGrid_Point = SeqNullSearch(vasID, Trim(.Text), 5)
            If pGrid_Point = 0 Then vasID.MaxRows = vasID.MaxRows + 1: pGrid_Point = vasID.MaxRows
            .RowHeight(-1) = 12
        End If
        
        .Row = Row: .Col = 4
        sBarcode = Trim(.Text)
        
        Call vasID.SetText(1, pGrid_Point, "1")
        Call vasID.SetText(4, pGrid_Point, .Text)
'        .Row = Row: .Col = 5
'        Call vasID.SetText(5, pGrid_Point, .Text)
'        .Row = Row: .Col = 6
'        Call vasID.SetText(6, pGrid_Point, .Text)
'        .Row = Row: .Col = 7
'        Call vasID.SetText(7, pGrid_Point, .Text)
'        .Row = Row: .Col = 8
'        Call vasID.SetText(8, pGrid_Point, .Text)
'        vasID.RowHeight(-1) = 12
    
        '¹ÙÄÚµå¹øÈ£·Î È¯ÀÚÁ¤º¸ ºÒ·¯¿À±â
              SQL = "SELECT DiSTINCT CHARTNO, PATNAME, PATSEX, PATAGE,COMPANY,HOSPCODE,PATJUMIN,PATNO,COMMDATE,EXAMNO,EXAMID,IOFLAG  "
        SQL = SQL & vbCrLf & "  FROM PAT_RES "
        SQL = SQL & vbCrLf & " WHERE EXAMTYPE = 'C' "
        SQL = SQL & vbCrLf & "   AND BARCODE = '" & sBarcode & "'"
        
    
        Res = GetDBSelectColumn(gLocal, SQL)
            
        If Res = 1 Then
            SetText frmInterface.vasID, Trim(gReadBuf(0)), pGrid_Point, colPID    '5
            SetText frmInterface.vasID, Trim(gReadBuf(1)), pGrid_Point, colPName  '6
            SetText frmInterface.vasID, Trim(gReadBuf(2)), pGrid_Point, colSex    '7
            SetText frmInterface.vasID, Trim(gReadBuf(3)), pGrid_Point, colAge    '8
            SetText frmInterface.vasID, Format(Trim(gReadBuf(8)), "####-##-##"), pGrid_Point, 2
            
            SetText frmInterface.vasID, Trim(gReadBuf(4)), pGrid_Point, 12
            SetText frmInterface.vasID, Trim(gReadBuf(5)), pGrid_Point, 13
            SetText frmInterface.vasID, Trim(gReadBuf(6)), pGrid_Point, 14
            SetText frmInterface.vasID, Trim(gReadBuf(7)), pGrid_Point, 15
            SetText frmInterface.vasID, Trim(gReadBuf(8)), pGrid_Point, 16
            SetText frmInterface.vasID, Trim(gReadBuf(9)), pGrid_Point, 17
            SetText frmInterface.vasID, Trim(gReadBuf(10)), pGrid_Point, 18
            SetText frmInterface.vasID, Trim(gReadBuf(11)), pGrid_Point, 19
            vasID.RowHeight(-1) = 12
        End If
    
    End With
    
End Sub

'Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'
'    If Winsock1.State <> sckClosed Then
'        Winsock1.Close
'
'        Winsock1.Accept requestID
'        StatusBar1.Panels(2).Text = "Àåºñ¿¡ Á¢¼ÓµÇ¾ú½À´Ï´Ù."
'    End If
'
'End Sub
'
'Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'    Dim strText As String
'    Dim strTmp As String
'
'    Dim strLastSeq  As String
'    Dim strRcvSign  As String
'    Dim strSendAck  As String
'    Dim strRcvCnt   As String
'
'    Dim strNS       As String
'    Dim strNE       As String
'    Dim intNS       As Integer
'    Dim intNE       As Integer
'
'    Dim strSendData  As String
'    Dim varBuffers   As Variant
'    Dim i As Integer
'    Dim lngBufLen As Long
'    Dim BufChar     As String
'
'    Winsock1.GetData strText
'
'        strBuffer = strText
'
'    SetRawData "[Rx]" & strBuffer
'    StatusBar1.Panels(3).Text = strBuffer
'
'    strBuffer = Replace(strBuffer, vbLf, "")
'
'    strRecvData = Split(strBuffer, vbCr)
'
'    Call EditRcvData
'
'
'End Sub

Private Sub vasWorkList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    Dim lsNM As String
    Dim lsPid As String
    Dim i As Integer
    Dim lsDate As String

    iRow = vasWorkList.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasWorkList.DataRowCnt Then
            Exit Sub
        End If

        lsID = Trim(GetText(vasWorkList, iRow, 4))
        lsPid = Trim(GetText(vasWorkList, iRow, colPID))
        lsNM = Trim(GetText(vasWorkList, iRow, colPName))
        lsDate = Trim(GetText(vasWorkList, iRow, 2))
        
        If MsgBox(lsNM & "À» Áö¿ì½Ã°Ú½À´Ï±î?", vbInformation + vbYesNo, "¾Ë¸²") = vbNo Then
            Exit Sub
        End If

        SQL = ""
        SQL = SQL & "DELETE FROM PAT_RES "
        SQL = SQL & " Where ChartNo  = '" & lsPid & "' "
        SQL = SQL & "   and CommDate = '" & Format(lsDate, "yyyymmdd") & "'"
        SQL = SQL & "   and BarCode  = '" & lsID & "'"
        SQL = SQL & "   and ExamType = 'C'"

        Res = SendQuery(gLocal, SQL)
        If Res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
        
        DeleteRow vasWorkList, iRow, iRow
        vasRes.MaxRows = 0

        vasWorkList.MaxRows = vasWorkList.MaxRows - 1
        
    End If
End Sub

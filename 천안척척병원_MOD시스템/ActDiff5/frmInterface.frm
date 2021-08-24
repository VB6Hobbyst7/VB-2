VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInterface 
   Caption         =   "ActDiff5"
   ClientHeight    =   11040
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   25560
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   11040
   ScaleWidth      =   25560
   Begin VB.PictureBox Picture1 
      Align           =   1  'À§ ¸ÂÃã
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   25500
      TabIndex        =   37
      Top             =   0
      Width           =   25560
      Begin VB.Label Label1 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "ActDiff5"
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
         Index           =   2
         Left            =   210
         TabIndex        =   41
         Top             =   90
         Width           =   735
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
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Port"
         Height          =   195
         Index           =   0
         Left            =   11640
         TabIndex        =   40
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Send"
         Height          =   195
         Left            =   12765
         TabIndex        =   39
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Receive"
         Height          =   195
         Left            =   13800
         TabIndex        =   38
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   3375
      Left            =   15180
      TabIndex        =   30
      Top             =   7230
      Visible         =   0   'False
      Width           =   8625
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1545
         Left            =   240
         TabIndex        =   31
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
         TabIndex        =   32
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
      TabIndex        =   8
      Top             =   660
      Visible         =   0   'False
      Width           =   8655
      Begin FPSpread.vaSpread vasCode 
         Height          =   1455
         Left            =   120
         TabIndex        =   22
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
         SpreadDesigner  =   "frmInterface.frx":4240
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   2235
         Left            =   3780
         TabIndex        =   9
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
         SpreadDesigner  =   "frmInterface.frx":4466
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         Picture         =   "frmInterface.frx":468C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   23
         Top             =   5790
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   210
         TabIndex        =   21
         Top             =   5640
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
         Height          =   585
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  '¼öÁ÷
         TabIndex        =   15
         Top             =   4830
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   5100
         TabIndex        =   14
         Top             =   5700
         Width           =   645
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
         Height          =   435
         Left            =   4440
         TabIndex        =   13
         Top             =   5715
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   3780
         MultiLine       =   -1  'True
         ScrollBars      =   3  '¾ç¹æÇâ
         TabIndex        =   12
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
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   11
         Top             =   5640
         Value           =   1  'È®ÀÎ
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   4860
         TabIndex        =   10
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
                  Picture         =   "frmInterface.frx":4C16
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":51B0
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":574A
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":5CE4
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":6576
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":66D0
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":682A
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1485
         Left            =   120
         TabIndex        =   16
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
         SpreadDesigner  =   "frmInterface.frx":6984
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   2205
         Left            =   3780
         TabIndex        =   17
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
         SpreadDesigner  =   "frmInterface.frx":6BAA
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1485
         Left            =   120
         TabIndex        =   18
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
         SpreadDesigner  =   "frmInterface.frx":6DD0
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'Åõ¸í
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
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
         Height          =   375
         Left            =   2010
         TabIndex        =   33
         Top             =   5760
         Width           =   825
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2940
         TabIndex        =   20
         Top             =   5730
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3720
         TabIndex        =   19
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
      TabPicture(0)   =   "frmInterface.frx":6FF6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtTest"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "ÀÌÀü°á°ú"
      TabPicture(1)   =   "frmInterface.frx":7012
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   9675
         Left            =   -74820
         TabIndex        =   58
         Top             =   360
         Width           =   14625
         Begin VB.OptionButton optSaveResultR 
            Caption         =   "¼öÁ¤"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
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
            TabIndex        =   72
            Top             =   270
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton optSaveResultR 
            Caption         =   "Àåºñ"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
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
            TabIndex        =   71
            Top             =   270
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   735
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
            Height          =   435
            Left            =   6810
            TabIndex        =   70
            Top             =   210
            Value           =   1  'È®ÀÎ
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   7860
            TabIndex        =   64
            Top             =   630
            Width           =   6675
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   69
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   1
               Left            =   4590
               TabIndex        =   68
               Top             =   240
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
               Left            =   3540
               TabIndex        =   67
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   1
               Left            =   1995
               TabIndex        =   66
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
               Left            =   510
               TabIndex        =   65
               Top             =   240
               Width           =   1380
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
            Left            =   13020
            TabIndex        =   63
            Top             =   -90
            Visible         =   0   'False
            Width           =   1395
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
            Left            =   3750
            TabIndex        =   62
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   61
            Top             =   780
            Width           =   225
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
            Left            =   5250
            TabIndex        =   60
            Top             =   240
            Visible         =   0   'False
            Width           =   1395
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
            TabIndex        =   59
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   8070
            Left            =   7860
            TabIndex        =   73
            Top             =   1455
            Width           =   6675
            _Version        =   393216
            _ExtentX        =   11774
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
            MaxCols         =   7
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":702E
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   74
            Top             =   300
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
            Format          =   21364736
            CurrentDate     =   40457
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   8805
            Left            =   150
            TabIndex        =   77
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
               Name            =   "±¼¸²Ã¼"
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
            SpreadDesigner  =   "frmInterface.frx":AD4A
            UserResize      =   2
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "°á°úÀû¿ë"
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
            Left            =   7890
            TabIndex        =   76
            Top             =   360
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "ÀÇ·ÚÀÏÀÚ"
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
            TabIndex        =   75
            Top             =   360
            Width           =   780
         End
      End
      Begin VB.TextBox txtTest 
         Height          =   375
         Left            =   3900
         TabIndex        =   43
         Top             =   30
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Àü¼ÛÅ×½ºÆ®"
         Height          =   435
         Left            =   4590
         TabIndex        =   42
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
         Begin FPSpread.vaSpread vasID 
            Height          =   8805
            Left            =   90
            TabIndex        =   79
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
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   15
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":B7D4
            UserResize      =   2
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
            Left            =   4680
            TabIndex        =   52
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cboChk 
            Height          =   315
            ItemData        =   "frmInterface.frx":C37D
            Left            =   3720
            List            =   "frmInterface.frx":C38A
            TabIndex        =   51
            Top             =   -30
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CommandButton cmdDownload 
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
            Left            =   5700
            TabIndex        =   50
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdPatDelete 
            Caption         =   "Á¦¿Ü"
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
            Left            =   6720
            TabIndex        =   49
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtNum 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   4140
            TabIndex        =   48
            Text            =   "0"
            Top             =   300
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.CommandButton cmdWorkDelete 
            Caption         =   "Á¦¿Ü"
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
            Left            =   6660
            TabIndex        =   47
            Top             =   4680
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox chkWAll 
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   420
            TabIndex        =   44
            Top             =   780
            Width           =   225
         End
         Begin VB.OptionButton optSaveResult 
            Caption         =   "¼öÁ¤"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
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
            TabIndex        =   35
            Top             =   -120
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton optSaveResult 
            Caption         =   "Àåºñ"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
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
            TabIndex        =   34
            Top             =   -120
            Visible         =   0   'False
            Width           =   735
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
            Left            =   11940
            TabIndex        =   7
            Top             =   240
            Width           =   1275
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
            Left            =   13260
            TabIndex        =   6
            Top             =   240
            Width           =   1245
         End
         Begin VB.Frame Frame6 
            Height          =   585
            Left            =   7860
            TabIndex        =   24
            Top             =   630
            Width           =   6675
            Begin VB.Label Label8 
               Caption         =   "µî·Ï¹øÈ£ :"
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
               Left            =   510
               TabIndex        =   29
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   0
               Left            =   1995
               TabIndex        =   28
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
               Left            =   3540
               TabIndex        =   27
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   0
               Left            =   4590
               TabIndex        =   26
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   25
               Top             =   720
               Width           =   1155
            End
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   450
            TabIndex        =   5
            Top             =   5160
            Width           =   225
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
            Left            =   1080
            TabIndex        =   45
            Top             =   330
            Width           =   2595
            _ExtentX        =   4577
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
            Format          =   21364736
            CurrentDate     =   40457
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   315
            Left            =   2580
            TabIndex        =   53
            Top             =   -90
            Visible         =   0   'False
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
            Format          =   21364737
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   315
            Left            =   1080
            TabIndex        =   54
            Top             =   -90
            Visible         =   0   'False
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
            Format          =   21364737
            CurrentDate     =   40248
         End
         Begin FPSpread.vaSpread vasWorkList 
            Height          =   3855
            Left            =   90
            TabIndex        =   55
            Top             =   720
            Visible         =   0   'False
            Width           =   7605
            _Version        =   393216
            _ExtentX        =   13414
            _ExtentY        =   6800
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
            MaxCols         =   15
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":C3A0
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8070
            Left            =   7770
            TabIndex        =   78
            Top             =   1440
            Width           =   6795
            _Version        =   393216
            _ExtentX        =   11986
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
            MaxCols         =   8
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":CF49
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Ã³¹æÀÏÀÚ"
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
            Index           =   3
            Left            =   180
            TabIndex        =   57
            Top             =   -30
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label12 
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
            Height          =   195
            Left            =   2400
            TabIndex        =   56
            Top             =   -30
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label Label1 
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
            Index           =   1
            Left            =   150
            TabIndex        =   46
            Top             =   390
            Width           =   780
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Æò¸é
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Åõ¸í
            Caption         =   "°á°úÀû¿ë"
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
            Left            =   7905
            TabIndex        =   36
            Top             =   -30
            Visible         =   0   'False
            Width           =   780
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '¾Æ·¡ ¸ÂÃã
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   10635
      Width           =   25560
      _ExtentX        =   45085
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
            TextSave        =   "2014-07-23"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "¿ÀÀü 11:28"
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
Const colSeqNo = 2
Const colPID = 3        'Áø·á¹øÈ£
Const colOrdDate = 4    '°Ë»çÀÇ·ÚÀÏ
Const colBarcode = 5    '¹ÙÄÚµå
Const colRack = 6       'Rack
Const colPos = 7        'Pos
Const colDept = 8       'Áø·á°ú
Const colIOFlag = 9     'ÀÔ¿ø/¿Ü·¡±¸ºÐ
Const colOrdSeq = 10    '¼ø¹ø
Const colReceNo = 11    'Á¢¼ö¹øÈ£(°Ë»çÀÇ·Ú)
Const colPName = 12     'È¯ÀÚ¸í
Const colOCnt = 13
Const colRCnt = 14
Const colState = 15


'seq Áø·á¹øÈ£ °Ë»çÀÇ·ÚÀÏ ¹ÙÄÚµå Áø·á°ú ±¸ºÐ ¼ø¹ø Á¢¼ö¹øÈ£ È¯ÀÚ¸í o r »óÅ×

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
                
                SetText vasID, Trim(GetText(vasWorkList, intRow, colSeqNo)), vasID.MaxRows, colSeqNo
                SetText vasID, Trim(GetText(vasWorkList, intRow, colOrdDate)), vasID.MaxRows, colOrdDate
                SetText vasID, Trim(GetText(vasWorkList, intRow, colBarcode)), vasID.MaxRows, colBarcode
                SetText vasID, Trim(GetText(vasWorkList, intRow, colPID)), vasID.MaxRows, colPID
                SetText vasID, Trim(GetText(vasWorkList, intRow, colPName)), vasID.MaxRows, colPName
                
'                SetText vasID, Trim(GetText(vasWorkList, intRow, colSex)), vasID.MaxRows, colSex
'                SetText vasID, Trim(GetText(vasWorkList, intRow, colAge)), vasID.MaxRows, colAge



                
''''                .Col = colBarcode
''''                SetText vasID, txtNum, vasID.MaxRows, colSeqNo
''''                SetText vasID, Trim(.Text), vasID.MaxRows, colBarcode
''''
''''                GetText vasID, vasID.MaxRows, colSeqNo
'''''                Call GetSampleInfoW(vasID.MaxRows)                                '5,6,7,8
''''
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
'            SetText vasPrint, Trim(GetText(vasRID, iRow, colSex)), j, 4
            
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
        MsgBox "ÀúÀåÇÒ ÀÚ·á°¡ ¾ø½À´Ï´Ù.", , "¾Ë ¸²"
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
    
    'SELECT Ã³À½ '' ´Â Ã¼Å©¹Ú½º
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
            Case "0": SetText vasRID, "¿¡·¯", iRow, colState
            Case "1": SetText vasRID, "°á°ú", iRow, colState
            Case "2": SetText vasRID, "¿Ï·á", iRow, colState
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
    Dim strDate As String
    
    vasWorkList.MaxRows = 0
    intRow = 0
        
          SQL = "Select Distinct a.PID, a.REQDATE, c.SPECIMENID, a.REQDEPT, a.IOFLAG, a.SEQNO, a.RECENO, b.PAT_NM " & vbCr
    SQL = SQL & "  From EXAMREQ a, TI_PAT b, EXAMRES c" & vbCr
    SQL = SQL & " Where a.PID = b.PAT_CHART" & vbCr
    SQL = SQL & "   And a.REQDATE Between '" & pFrDt & "' And '" & pToDt & "'" & vbCr
    SQL = SQL & "   And c.EXAMCODE in (" & gAllExam & ")" & vbCr
    SQL = SQL & "   And a.PID = c.PID "
    SQL = SQL & "   And a.SEQNO = c.SEQNO "
    SQL = SQL & "   And a.RECENO = c.RECENO "
    '-- °Ë»ç¿Ï·á
    SQL = SQL & "   And (c.EXAMEND = '' Or c.EXAMEND IS NULL) "
    
    SQL = SQL & " Order By a.REQDATE, a.REQDEPT, a.SEQNO, a.RECENO"
    Set RS = cn_Ser.Execute(SQL, , 1)

    Do Until RS.EOF
        intRow = intRow + 1
        vasWorkList.MaxRows = intRow
        
        SetText vasWorkList, "1", intRow, colCheckBox
        SetText vasWorkList, CStr(intRow), intRow, colSeqNo
        SetText vasWorkList, Trim(RS.Fields("PID")) & "", intRow, colPID
        SetText vasWorkList, Trim(RS.Fields("REQDATE")) & "", intRow, colOrdDate
        SetText vasWorkList, Trim(RS.Fields("SPECIMENID")) & "", intRow, colBarcode
        SetText vasWorkList, Trim(RS.Fields("REQDEPT")), intRow, colDept
        'SetText vasWorkList, Trim(RS.Fields("IOFLAG")), intRow, colIOFlag
        Select Case Trim(RS.Fields("IOFLAG"))
            Case "1": SetText vasWorkList, "¿Ü·¡", intRow, colIOFlag
            Case "2": SetText vasWorkList, "ÀÔ¿ø", intRow, colIOFlag
            'Case "I": SetText vasWorkList, "ÀÔ¿ø", intRow, colIOFlag
        End Select
        
        SetText vasWorkList, Trim(RS.Fields("SEQNO")), intRow, colOrdSeq
        SetText vasWorkList, Trim(RS.Fields("RECENO")), intRow, colReceNo
        SetText vasWorkList, Trim(RS.Fields("PAT_NM")), intRow, colPName
        
        RS.MoveNext
    Loop
    
    vasWorkList.RowHeight(-1) = 12

End Sub

Private Sub cmdSearch_Click()
                
    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
'    Call GetWorkList(dtpStartDt.Value, dtpStopDt.Value)
    
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

Private Sub imgPort_DblClick()
    
'    '-- °³¹ß½Ã¿¡¸¸ Remark Ç®¾î¼­ Å×½ºÆ®ÁøÇà
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
    strBuffer = strBuffer & "ÿ RESULT  " & vbCr
    strBuffer = strBuffer & "p 75" & vbCr
    strBuffer = strBuffer & "q 07/09/07 10h16mn07s" & vbCr
    strBuffer = strBuffer & "u 0000000000000001" & vbCr
    strBuffer = strBuffer & "s 0002" & vbCr
    strBuffer = strBuffer & "v                               " & vbCr
    strBuffer = strBuffer & "t R" & vbCr
    strBuffer = strBuffer & "€ D" & vbCr
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
'    strBuffer = strBuffer & "W           *.4027Faˆ²ÔðÿûêË°–†yiZQJFBB@DHKMJHBB?=94555222.****)))))))))*.000002799;@JOUUX\ejrwy{€†ˆ†Š“‘”“‘‘‘‹ˆ†ˆ?tpig_\XZ\ "
'    strBuffer=strBuffer & "X             !! !!!!!!""%&+-7>Pbw‘§ÄÌâõÿóñêÎÉ¶¦š?yl_WTICA=732.,+(++('''*%'''))''&&&%$&&#$$#"""""!"!!!       !                  '
'    strBuffer=strBuffer & "Y           #);KVW^aqs€|†…?†‚?uufciaaPLDGDC;;6610,-.210,,))&'''&&#$"$&&&$$""&&$! !!#"" !!! !""! !"!!!!""#!!!! """   !""""""$""!!
    strBuffer = strBuffer & "S       " & vbCr
    strBuffer = strBuffer & "_ 105" & vbCr
    strBuffer = strBuffer & "P           G3" & vbCr
    strBuffer = strBuffer & "] 000 000 000 027 039" & vbCr
    strBuffer = strBuffer & "?LC 250  " & vbCr
    strBuffer = strBuffer & "?V2.8 " & vbCr
    strBuffer = strBuffer & "?A75C" & vbCr
    strBuffer = strBuffer & "   " & vbCr

    strBuffer = STX & ";" & ETX & vbCr
    
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
        frmInterface.StatusBar1.Panels(2).Text = "¿¬°á µÇ¾ú½À´Ï´Ù"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        lblStatus = "ÀÛ¾÷Áß.."
    Else
        frmInterface.StatusBar1.Panels(2).Text = "¿¬°á µÇÁö ¾Ê¾Ò½À´Ï´Ù"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
'        lblStatus = "ÀÛ¾÷ ´ë±âÁß.."
    End If

    If Not Connect_Local Then
        MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    '-- osw Ãß°¡
    For i = 1 To 1
        If Not Connect_PRServer Then
            MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
            cn_Server_Flag = False
            Exit Sub
        Else
            cn_Server_Flag = True
        End If
    Next
    
    GetExamCode
    
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
    
    '-- ASTM TYPEº° Define ÇØ¾ßÇÔ.
    '-- ASTM TYPE = Standard
'1H|\^&|||BCI|||||||P|D1394-97|20140620091112
'42
'2P|1||||^|||U
'56
'3O|1|140619000012^02^01||^^^DIF|||||||||||||||||||||F
'E9

    
    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|\^&||||||||||||" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1||" & mOrder.BarNo & "||^|||U|||||Physician||||||||||||Location|||||||||" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
            
        Case 3  '## Order
           ' sSendBuff = m_iFrameN & "O|1|" & pSampleInfo.ID & "||" & "^^^" & pSampleInfo.IFCD(1) & "|||" & Format(Now, "yyyyMMddHHmmss") & "||||||||||||||||||O|||||" & vbCr & Chr(3)
            strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "||" & "^^^" & mOrder.Order & "|||" & Format(Now, "yyyyMMddHHmmss") & "||||||||||||||||||O|||||" & vbCr & ETX
            intSndPhase = 4
            intFrameNo = intFrameNo + 1
            
        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1
            
        Case 5  '## EOT
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
    
    strSndMsg = ">"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & vbCrLf
    
End Sub

Public Sub SndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) & GetChkSum(strSndMsg)
    comEqp.Output = strSndMsg & Chr(13)
    
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
'            Buffer = strBuffer
            
            'txtData = txtData & Buffer
            
            SetRawData "[Rx]" & Buffer
            lngBufLen = Len(Buffer)
            
            Debug.Print Buffer

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
                                        '-- Àåºñ¿¡¼­ ³Ñ¾î¿Â ½Ã°£ÀÌ ¿ì¿¬È÷ 11:59:59ÃÊ³ª ÀÍÀÏ¿¡ °¡±î¿î ½Ã°£ÀÏ °æ¿ì
                                        '-- °á°ú ÀúÀå½Ã ÀÌÀüÀÏÀ» °¡Á®¿Ã ¼ö ÀÖÀ¸¹Ç·Î ³¯Â¥¸¦ ½Ç½Ã°£ ¾÷µ¥ÀÌÆ® ÇÑ´Ù.
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
                                    Case vbCr
                                    Case vbLf
                                        intPhase = 4
                                        comEqp.Output = ACK
                                        SetRawData "[Tx]" & ACK
                                End Select
                            Case 4      '## Termination Phase
                                Select Case BufChar
                                    Case STX
                                        intPhase = 2
                                    Case EOT
                                        '-- Àåºñ¿¡¼­ ³Ñ¾î¿Â ½Ã°£ÀÌ ¿ì¿¬È÷ 11:59:59ÃÊ³ª ÀÍÀÏ¿¡ °¡±î¿î ½Ã°£ÀÏ °æ¿ì
                                        '-- °á°ú ÀúÀå½Ã ÀÌÀüÀÏÀ» °¡Á®¿Ã ¼ö ÀÖÀ¸¹Ç·Î ³¯Â¥¸¦ ½Ç½Ã°£ ¾÷µ¥ÀÌÆ® ÇÑ´Ù.
                                        strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                        dtpToday.Value = Format(strDate, "####-##-##")
                                        
                                        DoEvents
                                        
                                        Call EditRcvDataASTM
                                        If strState = "Q" Then
                                            intSndPhase = 1
                                            If gCOMFormat = "1" Then
                                                intFrameNo = 0
                                            Else 'If gComFormat = "2" Then
                                                intFrameNo = 1
                                            End If
                                            comEqp.Output = ENQ
                                            SetRawData "[Tx]" & ENQ
                                        End If
                                        
                                        intPhase = 1
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

'-----------------------------------------------------------------------------'
'   ±â´É : ÇØ´ç ¹ÙÄÚµå¹øÈ£¿¡ ´ëÇÑ Á¢¼öÁ¤º¸ Á¶È¸, Ç¥½Ã, °Ë»ç¿À´õ¸¸µé±â
'   ÀÎ¼ö :
'       - pBarNo : ¹ÙÄÚµå¹øÈ£
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
    End If
    
    '-- Àåºñ¼ö½ÅÁ¤º¸ Ç¥½Ã
    Call SetText(vasID, pBarNo, intRow, colBarcode)         '3  ¹ÙÄÚµå
    Call SetText(vasID, mOrder.RackNo, intRow, colRack)     '4  Rack¹øÈ£
    Call SetText(vasID, mOrder.TubePos, intRow, colPos)     '5  Pos¹øÈ£
    
    Call vasActiveCell(vasID, intRow, colBarcode)
    Call ClearSpread(vasRes)
    
    '-- °Ë»çÀÚ Á¤º¸ ¼­¹öÅ×ÀÌºí¿¡¼­ °¡Á®¿Í Ç¥½Ã(for ¿öÅ©¸®½ºÆ®)  '6,7,8,9
    Call GetSampleInfoW(intRow)
    
    '-- ¹ÙÄÚµå¹øÈ£¿¡ Á¸ÀçÇÏ´Â °Ë»çÄÚµå °¡Á®¿À±â(ÀÎ¼ö : ÀåºñÄÚµå,¹ÙÄÚµå¹øÈ£)
    gOrderExam = GetOrderExamCode(gEquip, pBarNo)
    
    '-- ·ÎÄÃÅ×ÀÌºí¿¡¼­ °Ë»çÇ×¸ñ¿¡ ÇØ´çÇÏ´Â °Ë»çÃ¤³Î Ã£¾Æ¿À±â (intRow = ±âÁ¸ °Ë»çÇß´ø ¹ÙÄÚµå°¡ ´Ù½Ã ¿Ã¶ó¿Ã °æ¿ì À§Ä¡¸¦ ¸øÃ£´Â´Ù.)
    strItems = GetGetEquipExamCode_ActDiff5AL(gEquip, pBarNo, intRow)

    '-- °Ë»çÃ¤³Î·Î Àåºñ¿À´õ ¸¸µé±â
    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = strItems
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
    End If
    
    Call SetText(vasID, "Order", intRow, colState)         '12 ÁøÇà»óÅÂ

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
    Dim strOrdDt    As String
    Dim strTestDt   As String
    
    
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
    
    '2014-12-30
    strTestDt = Mid(dtpToday.Value, 1, 4) & Mid(dtpToday.Value, 6, 2) & Mid(dtpToday.Value, 9, 2)
    
    
    '-- Àåºñ¼ö½ÅÁ¤º¸ Ç¥½Ã
    Call SetText(vasID, strTestDt, intRow, colOrdDate)
    Call SetText(vasID, pBarNo, intRow, colBarcode)
    Call SetText(vasID, mResult.RackNo, intRow, colRack)
    Call SetText(vasID, mResult.TubePos, intRow, colPos)
    Call vasActiveCell(vasID, intRow, colBarcode)
    
    '-- °á°ú½ºÇÁ·¹µå Áö¿ì±â
    Call ClearSpread(vasRes)
    
    '-- °Ë»çÀÚ Á¤º¸ ¼­¹öÅ×ÀÌºí °¡Á®¿Í Ç¥½Ã(for ¿öÅ©¸®½ºÆ®)  '5,6,7,8
    Call GetSampleInfoW(intRow)                                '5,6,7,8
    
    '-- ÇöÀç Row
    gRow = intRow
    
    '-- ¹ÙÄÚµå¹øÈ£¿¡ Á¸ÀçÇÏ´Â °Ë»çÄÚµå °¡Á®¿À±â(ÀÎ¼ö : ÀåºñÄÚµå,¹ÙÄÚµå¹øÈ£)
    gOrderExam = GetOrderExamCode(gEquip, pBarNo, strOrdDt)

    
End Sub



'-----------------------------------------------------------------------------'
'   ±â´É : Àåºñ·ÎºÎ ¼ö½ÅÇÑ µ¥ÀÌÅÍ ÆíÁý
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM()
    Dim strRcvBuf    As String   '¼ö½ÅÇÑ Data
    Dim strType      As String   '¼ö½ÅÇÑ Record Type
    Dim strBarno     As String   '¼ö½ÅÇÑ ¹ÙÄÚµå¹øÈ£
    Dim strSeq       As String   '¼ö½ÅÇÑ Sequence
    Dim strRackNo    As String   '¼ö½ÅÇÑ Rack Or Disk No
    Dim strTubePos   As String   '¼ö½ÅÇÑ Tube Position
    Dim strIntBase   As String   '¼ö½ÅÇÑ Àåºñ±âÁØ °Ë»ç¸í
    Dim strResult    As String   '¼ö½ÅÇÑ °á°ú(Á¤¼º)
    Dim strIntResult As String   '¼ö½ÅÇÑ °á°ú(Á¤·®)
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
    Dim varRcvBuf   As Variant
    Dim intRow      As Integer
    Dim i As Integer
    
    
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        
        SetRawData "[strRcvBuf]" & strRcvBuf
'2Q|1|140017856||||||||||NB2
    
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "Q"
                strBarno = Trim(mGetP(strRcvBuf, 3, "|"))
                
                SetRawData "[strBarno]" & strBarno
                
                If strBarno = "" Then Exit Sub
                
                With mOrder
                    .BarNo = strBarno
                End With
                
                strBarno = strBarno & "1"
                
                'Call SetPatInfoQry(strBarno)
                'MsgBox "1"
                
                
                Call GetOrder(strBarno)
                 
                strState = "Q"
                
            Case "O"    '## Order
                strTemp1 = Trim(mGetP(strRcvBuf, 3, "|"))
                strBarno = Trim(mGetP(strTemp1, 1, "^"))
                strRackNo = Trim(mGetP(strTemp1, 2, "^"))
                strTubePos = Trim(mGetP(strTemp1, 3, "^"))
                
                
                If strBarno = "" Then Exit Sub
                If Len(strBarno) <> 12 Then Exit Sub
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                End With
                
                Call SetPatInfo(strBarno)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strState = "O"
                
            Case "R"    '## Result
                '## Àåºñ±âÁØ °Ë»ç¸í, °á°ú, Abnormal Flag
                strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
                strResult = mGetP(strRcvBuf, 4, "|")
                
                If strResult <> "" Then
                    SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                    SQL = SQL & "  FROM EQPMASTER"
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
                        SetText vasID, "Result", gRow, colState                 '11 ÁøÇà»óÅÂ
                        
                        '-- °á°ú List
                        SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                        SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                        SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     'Àåºñ°á°ú
                        SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                        SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                        '-- ·ÎÄÃ ÀúÀå
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                        
                    '-- ¿À´õ ¾øÀ» °æ¿ì
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
                            
                            '¼Ò¼öÁ¡ Ã³¸®, °á°ú ÇüÅÂ Ã³¸®
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText vasID, "Result", gRow, colState                 'ÁøÇà»óÅÂ
                            
                            '-- °á°ú List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                            SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     'Àåºñ°á°ú
                            SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- ·ÎÄÃ ÀúÀå
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            'strState = ""
                        End If
                    End If
                End If
            Case "C"    '## Comment
                '## Abnormal °á°úÀÏ¶§ Comment ÀúÀå
            Case "L"    '## Terminator
                '## DB¿¡ °á°úÀúÀå
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
                        SetText vasID, "0", gRow, colCheckBox
                        
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
                    
                    strState = ""
                
                End If
            
                'SetText vasID, "Result", gRow, colState
        
        End Select
    Next

End Sub

'-----------------------------------------------------------------------------'
'   ±â´É : Àåºñ·ÎºÎ ¼ö½ÅÇÑ µ¥ÀÌÅÍ ÆíÁý
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataGAS()
    Dim strRcvBuf    As String   '¼ö½ÅÇÑ Data
    Dim strType      As String   '¼ö½ÅÇÑ Record Type
    Dim strBarno     As String   '¼ö½ÅÇÑ ¹ÙÄÚµå¹øÈ£
    Dim strSeq       As String   '¼ö½ÅÇÑ Sequence
    Dim strRackNo    As String   '¼ö½ÅÇÑ Rack Or Disk No
    Dim strTubePos   As String   '¼ö½ÅÇÑ Tube Position
    Dim strIntBase   As String   '¼ö½ÅÇÑ Àåºñ±âÁØ °Ë»ç¸í
    Dim strResult    As String   '¼ö½ÅÇÑ °á°ú(Á¤¼º)
    Dim strIntResult As String   '¼ö½ÅÇÑ °á°ú(Á¤·®)
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
                    If Trim(GetText(vasID, ii, colSeqNo)) = strSeq Then
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
                '## Àåºñ±âÁØ °Ë»ç¸í, °á°ú, Abnormal Flag
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
                        SetText vasID, "Result", gRow, colState                 '11 ÁøÇà»óÅÂ
                        

                        '-- °á°ú List
                        SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                        SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                        SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                        SetText vasRes, lsEquipRes, lsResRow, colMachResult     'Àåºñ°á°ú
                        SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                        SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                        SetText vasRes, strComm, lsResRow, 7                    'Flag
                        '-- ·ÎÄÃ ÀúÀå
                        SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                        
                    '-- ¿À´õ ¾øÀ» °æ¿ì
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
                            
                            '¼Ò¼öÁ¡ Ã³¸®, °á°ú ÇüÅÂ Ã³¸®
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText vasID, "Result", gRow, colState                 'ÁøÇà»óÅÂ
                            
                            '-- °á°ú List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                            SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     'Àåºñ°á°ú
                            SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                            SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                            '-- ·ÎÄÃ ÀúÀå
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            
                        End If
                    End If
                End If
            Case 50   '## Terminator
                '## DB¿¡ °á°úÀúÀå
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
'   ±â´É : Àåºñ·ÎºÎ ¼ö½ÅÇÑ µ¥ÀÌÅÍ ÆíÁý
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataAU()
    Dim strRcvBuf    As String   '¼ö½ÅÇÑ Data
    Dim strType      As String   '¼ö½ÅÇÑ Record Type
    Dim strBarno     As String   '¼ö½ÅÇÑ ¹ÙÄÚµå¹øÈ£
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
    
    
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        strType = Mid$(strRcvBuf, 1, 2)
        
        Select Case strType
            '## Order Begin =========================================
            Case "RB"   '## Begin Inquiry Text
            Case "R "    '## Inquiry Order
                strBarno = Trim(Mid(strRcvBuf, 14, 20))
                strRackNo = Mid(strRcvBuf, 3, 4)
                strTubePos = Mid(strRcvBuf, 7, 2)
                
                With mOrder
                    .BarNo = strBarno
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .Seq = Mid(strRcvBuf, 9, 5)
                End With
                
                Call GetOrder(strBarno)
                
            Case "RE"   '## End Inquirty Text
            
            '## Result =========================================
            Case "DB"   '## Begin Result Text
            Case "D "    '## Result
                strBarno = Trim$(Mid$(strRcvBuf, 14, 10))
                
                With mResult
                    .BarNo = strBarno
                    .RackNo = Mid(strRcvBuf, 3, 4)
                    .TubePos = Mid(strRcvBuf, 7, 2)
                End With
                
                If strBarno = "" Then Exit Sub

                strTmp = Mid$(strRcvBuf, 29)
                                
                Call SetPatInfo(strBarno)
                
                Do While Len(strTmp) >= 11
                    strIntBase = Mid$(strTmp, 2, 2)
                    strResult = Mid$(strTmp, 4, 6)
                    strComm = Mid$(strTmp, 10, 1)
        
                    If strResult <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
                        SQL = SQL & "  FROM EQPMASTER"
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
                            SetText vasID, "Result", gRow, colState                 '11 ÁøÇà»óÅÂ
                            

                            '-- °á°ú List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                            SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     'Àåºñ°á°ú
                            SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                            SetText vasRes, strComm, lsResRow, 7                    'Flag
                            '-- ·ÎÄÃ ÀúÀå
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            
                            strState = "R"
                            
                        '-- ¿À´õ ¾øÀ» °æ¿ì
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
                                
                                '¼Ò¼öÁ¡ Ã³¸®, °á°ú ÇüÅÂ Ã³¸®
                                lsEquipRes = strResult
                                strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                '-- Work List
                                SetText vasID, "Result", gRow, colState                 'ÁøÇà»óÅÂ
                                
                                '-- °á°ú List
                                SetText vasRes, strIntBase, lsResRow, colEquipCode      'ÀåºñÄÚµå
                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '°Ë»çÄÚµå
                                SetText vasRes, lsExamName, lsResRow, colExamName       '°Ë»ç¸í
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     'Àåºñ°á°ú
                                SetText vasRes, strResult, lsResRow, colResult          '°á°ú
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '¼ø¹ø
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                '-- ·ÎÄÃ ÀúÀå
                                SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                            
                                lsResult_Buff = ""
                                
                            End If
                        End If
                    End If
                    strTmp = Mid$(strTmp, 12)
                Loop
                
            
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
' asRow2 = °á°ú List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    
    'sExamDate = Format(dtpToday, "yyyymmdd")
    sExamDate = Trim(GetText(vasID, asRow1, colOrdDate))
    
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
                "PID,PNAME,RECENO,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
                "EQUIPRESULT,RESULT,EXAMNAME,SENDFLAG,EXAMUID) " & vbCrLf
    SQL = SQL & "VALUES("
'    SQL = SQL & "'" & Trim(Format(dtpExamDate.Value, "YYYYMMDD")) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colOrdDate)) & "', "
    SQL = SQL & "'" & gEquip & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colBarcode)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colRack)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPID)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colPName)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasID, asRow1, colReceNo)) & "', "
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
    sMsg = "°Ë»çÀÚ¸¦ ÀÔ·ÂÇØÁÖ¼¼¿ä."
    lblUser.Caption = InputBox(sMsg, "°Ë»çÀÚ ÀÔ·Â")

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
    'Local¿¡¼­ ºÒ·¯¿À±â
    ClearSpread vasRes
    
    'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
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
'        If MsgBox("ÇØ´ç È¯ÀÚ°á°ú¸¦ »èÁ¦ÇÏ½Ã°Ú½À´Ï±î?", vbInformation + vbYesNo, "¾Ë¸²") = vbNo Then
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
'        'Local¿¡¼­ ºÒ·¯¿À±â
'        ClearSpread vasTemp
'
'        'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
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
    'Local¿¡¼­ ºÒ·¯¿À±â
    ClearSpread vasRRes
    
    'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
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
'        'Local¿¡¼­ ºÒ·¯¿À±â
'        ClearSpread vasTemp
'
'        'ÀåºñÄÚµå, °Ë»çÄÚµå, °Ë»ç¸í, °á°ú, ¼ø¹ø
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
'        If MsgBox("ÇØ´ç È¯ÀÚ°á°ú¸¦ »èÁ¦ÇÏ½Ã°Ú½À´Ï±î?", vbInformation + vbYesNo, "¾Ë¸²") = vbNo Then
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
        txtNum.Text = txtNum.Text + 1
        
        .Col = colBarcode
        SetText vasID, txtNum, vasID.MaxRows, colSeqNo
        SetText vasID, Trim(.Text), vasID.MaxRows, colBarcode
'        Call GetSampleInfoW(vasID.MaxRows)

'        SetText vasID, Trim(GetText(vasWorkList, Row, colSeqNo)), vasID.MaxRows, colSeqNo
        SetText vasID, Trim(GetText(vasWorkList, Row, colOrdDate)), vasID.MaxRows, colOrdDate
        SetText vasID, Trim(GetText(vasWorkList, Row, colBarcode)), vasID.MaxRows, colBarcode
        SetText vasID, Trim(GetText(vasWorkList, Row, colPID)), vasID.MaxRows, colPID
        SetText vasID, Trim(GetText(vasWorkList, Row, colPName)), vasID.MaxRows, colPName


'        .Action = ActionDeleteRow
'        .MaxRows = .MaxRows - 1
    End With

End Sub


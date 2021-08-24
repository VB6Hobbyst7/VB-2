VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInterface 
   Caption         =   "  KRYPTOR Interface "
   ClientHeight    =   11160
   ClientLeft      =   345
   ClientTop       =   840
   ClientWidth     =   16755
   BeginProperty Font 
      Name            =   "����ü"
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
   ScaleHeight     =   11160
   ScaleWidth      =   16755
   StartUpPosition =   1  '������ ���
   Begin VB.Frame Frame4 
      Caption         =   "Print"
      Height          =   3375
      Left            =   16830
      TabIndex        =   21
      Top             =   7350
      Visible         =   0   'False
      Width           =   8625
      Begin FPSpread.vaSpread vasPrintBuf 
         Height          =   1245
         Left            =   240
         TabIndex        =   108
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
         SpreadDesigner  =   "frmInterface.frx":14F5
      End
      Begin FPSpread.vaSpread vasPrint 
         Height          =   1245
         Left            =   240
         TabIndex        =   109
         Top             =   1710
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
         SpreadDesigner  =   "frmInterface.frx":172E
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   6375
      Left            =   16800
      TabIndex        =   5
      Top             =   780
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame Frame2 
         Caption         =   "Error Log"
         Height          =   1815
         Left            =   780
         TabIndex        =   51
         Top             =   3420
         Visible         =   0   'False
         Width           =   5970
         Begin VB.TextBox txtErrLog 
            Appearance      =   0  '���
            Height          =   1455
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '����
            TabIndex        =   52
            Top             =   240
            Width           =   5775
         End
      End
      Begin FPSpread.vaSpread vasCode 
         Height          =   1455
         Left            =   120
         TabIndex        =   19
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
         SpreadDesigner  =   "frmInterface.frx":1967
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   2235
         Left            =   3780
         TabIndex        =   6
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
         SpreadDesigner  =   "frmInterface.frx":1BA0
      End
      Begin VB.PictureBox picLogin 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         Picture         =   "frmInterface.frx":1DD9
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   20
         Top             =   5790
         Width           =   285
      End
      Begin VB.CommandButton lblclear 
         Caption         =   "lblclear"
         Height          =   495
         Left            =   210
         TabIndex        =   18
         Top             =   5640
         Width           =   1215
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "����"
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
         ScrollBars      =   2  '����
         TabIndex        =   12
         Top             =   4830
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.TextBox txtTemp 
         Height          =   435
         Left            =   5100
         TabIndex        =   11
         Top             =   5700
         Width           =   645
      End
      Begin VB.TextBox Text_ini 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4440
         TabIndex        =   10
         Top             =   5715
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   3780
         MultiLine       =   -1  'True
         ScrollBars      =   3  '�����
         TabIndex        =   9
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
         Style           =   1  '�׷���
         TabIndex        =   8
         Top             =   5640
         Value           =   1  'Ȯ��
         Width           =   1065
      End
      Begin VB.Frame FrmUseControl 
         Caption         =   "UseControl"
         Height          =   870
         Left            =   4860
         TabIndex        =   7
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
                  Picture         =   "frmInterface.frx":2363
                  Key             =   "RUN"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":28FD
                  Key             =   "NOT"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":2E97
                  Key             =   "STOP"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":3431
                  Key             =   "LST"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":3CC3
                  Key             =   "ITM"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":3E1D
                  Key             =   "ERR"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInterface.frx":3F77
                  Key             =   "NOF"
               EndProperty
            EndProperty
         End
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1485
         Left            =   120
         TabIndex        =   13
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
         SpreadDesigner  =   "frmInterface.frx":40D1
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   2205
         Left            =   3780
         TabIndex        =   14
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
         SpreadDesigner  =   "frmInterface.frx":430A
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1485
         Left            =   120
         TabIndex        =   15
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
         SpreadDesigner  =   "frmInterface.frx":4543
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '����
         BorderStyle     =   1  '���� ����
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   22
         Top             =   5760
         Width           =   825
      End
      Begin VB.Label lblChangeBar 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   2940
         TabIndex        =   17
         Top             =   5730
         Width           =   735
      End
      Begin VB.Label lblChangePID 
         BackColor       =   &H000000FF&
         Height          =   405
         Left            =   3720
         TabIndex        =   16
         Top             =   5730
         Width           =   705
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   10185
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   16650
      _ExtentX        =   29369
      _ExtentY        =   17965
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "WorkList"
      TabPicture(0)   =   "frmInterface.frx":477C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtTest"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "�����ȸ"
      TabPicture(1)   =   "frmInterface.frx":4798
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkRAll"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chkRAll 
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   -74280
         TabIndex        =   93
         Top             =   1140
         Width           =   225
      End
      Begin VB.Frame Frame3 
         Height          =   9705
         Left            =   -74850
         TabIndex        =   73
         Top             =   360
         Width           =   16365
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   14790
            TabIndex        =   110
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   8895
            Left            =   120
            TabIndex        =   107
            Top             =   690
            Width           =   16095
            _Version        =   393216
            _ExtentX        =   28390
            _ExtentY        =   15690
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridShowHoriz   =   0   'False
            GridShowVert    =   0   'False
            MaxCols         =   26
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":47B4
            UserResize      =   2
         End
         Begin VB.CheckBox chkQC 
            Caption         =   "QC"
            Height          =   285
            Left            =   4200
            TabIndex        =   106
            Top             =   300
            Width           =   555
         End
         Begin VB.OptionButton optSaveResultR 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   10845
            TabIndex        =   86
            Top             =   270
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton optSaveResultR 
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   10065
            TabIndex        =   85
            Top             =   270
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chkSave 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   0
            Left            =   4800
            TabIndex        =   84
            Top             =   210
            Width           =   765
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   7860
            TabIndex        =   78
            Top             =   630
            Width           =   6675
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   83
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   1
               Left            =   4590
               TabIndex        =   82
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label Label4 
               Caption         =   "ȯ�ڸ� :"
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3540
               TabIndex        =   81
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   1
               Left            =   1995
               TabIndex        =   80
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label2 
               Caption         =   "���ڵ��ȣ :"
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   510
               TabIndex        =   79
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "&Excel"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   13290
            TabIndex        =   77
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdRSch 
            Caption         =   "�����ȸ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5610
            TabIndex        =   76
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdRTrans 
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7110
            TabIndex        =   75
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdRClear 
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11790
            TabIndex        =   74
            Top             =   240
            Width           =   1395
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   8070
            Left            =   7860
            TabIndex        =   87
            Top             =   1455
            Width           =   6675
            _Version        =   393216
            _ExtentX        =   11774
            _ExtentY        =   14235
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            SpreadDesigner  =   "frmInterface.frx":56E9
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   88
            Top             =   300
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   66191361
            CurrentDate     =   40457
         End
         Begin FPSpread.vaSpread vasRID1 
            Height          =   8805
            Left            =   165
            TabIndex        =   89
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
               Name            =   "����ü"
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
            SpreadDesigner  =   "frmInterface.frx":940C
            UserResize      =   2
         End
         Begin MSComCtl2.DTPicker dtpExamDateTo 
            Height          =   315
            Left            =   2700
            TabIndex        =   105
            Top             =   300
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   66191361
            CurrentDate     =   40457
         End
         Begin VB.Label Label7 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   8850
            TabIndex        =   91
            Top             =   360
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Label Label9 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�˻�����"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   90
            Top             =   390
            Width           =   780
         End
      End
      Begin VB.TextBox txtTest 
         Height          =   375
         Left            =   16080
         TabIndex        =   36
         Top             =   -30
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton Command16 
         Caption         =   "�����׽�Ʈ"
         Height          =   435
         Left            =   16770
         TabIndex        =   35
         Top             =   -60
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   9705
         Left            =   150
         TabIndex        =   2
         Top             =   360
         Width           =   16365
         Begin VB.CommandButton cmdAddRow 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   150
            TabIndex        =   104
            Top             =   720
            Width           =   375
         End
         Begin VB.Frame Frame6 
            Enabled         =   0   'False
            Height          =   525
            Left            =   10080
            TabIndex        =   99
            Top             =   -150
            Visible         =   0   'False
            Width           =   1455
            Begin VB.TextBox txtPos 
               Alignment       =   2  '��� ����
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   960
               TabIndex        =   101
               Text            =   "1"
               Top             =   150
               Width           =   375
            End
            Begin VB.TextBox txtCass 
               Alignment       =   2  '��� ����
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   300
               TabIndex        =   100
               Text            =   "A"
               Top             =   150
               Width           =   375
            End
            Begin VB.Label Label1 
               Appearance      =   0  '���
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  '����
               Caption         =   "P"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Index           =   5
               Left            =   750
               TabIndex        =   103
               Top             =   180
               Width           =   165
            End
            Begin VB.Label Label1 
               Appearance      =   0  '���
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  '����
               Caption         =   "C"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Index           =   4
               Left            =   60
               TabIndex        =   102
               Top             =   180
               Width           =   180
            End
         End
         Begin VB.CheckBox chkSave 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   1
            Left            =   5940
            TabIndex        =   98
            Top             =   210
            Width           =   735
         End
         Begin VB.Frame FraSearch 
            Appearance      =   0  '���
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   5580
            TabIndex        =   95
            Top             =   3360
            Visible         =   0   'False
            Width           =   4005
            Begin MSComctlLib.ProgressBar progBar 
               Height          =   195
               Left            =   210
               TabIndex        =   97
               Top             =   660
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   344
               _Version        =   393216
               Appearance      =   0
            End
            Begin VB.Label Label6 
               Appearance      =   0  '���
               BackColor       =   &H80000005&
               Caption         =   "������ ��ȸ���Դϴ�."
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   270
               TabIndex        =   96
               Top             =   300
               Width           =   3405
            End
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9870
            TabIndex        =   94
            Top             =   240
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdOrder 
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11640
            TabIndex        =   92
            Top             =   -30
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Frame FraLog 
            Caption         =   "Communication Log"
            Height          =   5355
            Left            =   60
            TabIndex        =   71
            Top             =   4260
            Visible         =   0   'False
            Width           =   8760
            Begin MSWinsockLib.Winsock Winsock1 
               Left            =   2250
               Top             =   540
               _ExtentX        =   741
               _ExtentY        =   741
               _Version        =   393216
            End
            Begin VB.Timer Timer1 
               Enabled         =   0   'False
               Interval        =   500
               Left            =   6210
               Top             =   660
            End
            Begin VB.TextBox txtComLog 
               Appearance      =   0  '���
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   8.25
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4995
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  '����
               TabIndex        =   72
               Top             =   240
               Width           =   8535
            End
         End
         Begin VB.CommandButton cmdIFClear 
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7800
            TabIndex        =   69
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtStopNum 
            Alignment       =   2  '��� ����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   5250
            TabIndex        =   68
            Top             =   270
            Width           =   555
         End
         Begin VB.CheckBox chkAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   570
            TabIndex        =   58
            Top             =   750
            Width           =   225
         End
         Begin VB.Frame Frame7 
            Height          =   570
            Left            =   7980
            TabIndex        =   42
            Top             =   2010
            Visible         =   0   'False
            Width           =   12465
            Begin VB.Label lblPos 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1110
               TabIndex        =   67
               Top             =   900
               Width           =   1905
            End
            Begin VB.Label lblOther 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3900
               TabIndex        =   66
               Top             =   900
               Width           =   7845
            End
            Begin VB.Label lblRack 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1110
               TabIndex        =   65
               Top             =   570
               Width           =   1905
            End
            Begin VB.Label lblSpcNm 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3900
               TabIndex        =   64
               Top             =   570
               Width           =   2205
            End
            Begin VB.Label lblSA 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7260
               TabIndex        =   63
               Top             =   570
               Width           =   1995
            End
            Begin VB.Label lblPatNm 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7260
               TabIndex        =   62
               Top             =   240
               Width           =   1995
            End
            Begin VB.Label lblCustNm 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3900
               TabIndex        =   61
               Top             =   240
               Width           =   2205
            End
            Begin VB.Label lblPtId 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  '����
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1110
               TabIndex        =   60
               Top             =   240
               Width           =   1905
            End
            Begin VB.Label lblControl 
               AutoSize        =   -1  'True
               Caption         =   "�Ƿڹ�ȣ :"
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   150
               TabIndex        =   50
               Top             =   240
               Width           =   900
            End
            Begin VB.Label lblLevel 
               AutoSize        =   -1  'True
               Caption         =   "Rack :"
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   150
               TabIndex        =   49
               Top             =   600
               Width           =   540
            End
            Begin VB.Label lblLotNo 
               AutoSize        =   -1  'True
               Caption         =   "Pos :"
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   150
               TabIndex        =   48
               Top             =   975
               Width           =   450
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               Caption         =   "�ŷ�ó :"
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   3105
               TabIndex        =   47
               Top             =   240
               Width           =   720
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               Caption         =   "��ü�� :"
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   3105
               TabIndex        =   46
               Top             =   600
               Width           =   720
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               Caption         =   "��Ÿ : "
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   3105
               TabIndex        =   45
               Top             =   975
               Width           =   630
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               Caption         =   "ȯ���̸� :"
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   3
               Left            =   6270
               TabIndex        =   44
               Top             =   240
               Width           =   900
            End
            Begin VB.Label lblGeneral 
               AutoSize        =   -1  'True
               Caption         =   "����/����:"
               BeginProperty Font 
                  Name            =   "����ü"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   4
               Left            =   6270
               TabIndex        =   43
               Top             =   600
               Width           =   900
            End
         End
         Begin FPSpread.vaSpread vasWorkList 
            Height          =   8985
            Left            =   120
            TabIndex        =   33
            Top             =   690
            Width           =   16095
            _Version        =   393216
            _ExtentX        =   28390
            _ExtentY        =   15849
            _StockProps     =   64
            ColHeaderDisplay=   0
            ColsFrozen      =   1
            EditModeReplace =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridShowHoriz   =   0   'False
            GridShowVert    =   0   'False
            MaxCols         =   26
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":9E44
            UserResize      =   2
         End
         Begin VB.TextBox txtStartNum 
            Alignment       =   2  '��� ����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   4500
            TabIndex        =   40
            Top             =   270
            Width           =   555
         End
         Begin VB.CheckBox chkWAll 
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   39
            Top             =   780
            Width           =   225
         End
         Begin VB.CommandButton cmdPatDelete 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8820
            TabIndex        =   38
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdDownload 
            Caption         =   "Down"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   14130
            TabIndex        =   37
            Top             =   30
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cboChk 
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            ItemData        =   "frmInterface.frx":AD79
            Left            =   17310
            List            =   "frmInterface.frx":AD83
            TabIndex        =   29
            Top             =   240
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "��ȸ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6780
            TabIndex        =   28
            Top             =   240
            Width           =   975
         End
         Begin FPSpread.vaSpread vasID 
            Height          =   4425
            Left            =   -510
            TabIndex        =   4
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
               Name            =   "����ü"
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
            SpreadDesigner  =   "frmInterface.frx":AD93
            UserResize      =   2
         End
         Begin FPSpread.vaSpread vasRes 
            Height          =   8850
            Left            =   8970
            TabIndex        =   3
            Top             =   2820
            Visible         =   0   'False
            Width           =   12465
            _Version        =   393216
            _ExtentX        =   21987
            _ExtentY        =   15610
            _StockProps     =   64
            BackColorStyle  =   1
            ColHeaderDisplay=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   14
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":B8BA
         End
         Begin MSComCtl2.DTPicker dtpStopDt 
            Height          =   315
            Left            =   2550
            TabIndex        =   30
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   66191361
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpStartDt 
            Height          =   315
            Left            =   1080
            TabIndex        =   31
            Top             =   270
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   66191361
            CurrentDate     =   40248
         End
         Begin FPSpread.vaSpread tblErrors 
            Height          =   1515
            Left            =   8850
            TabIndex        =   53
            Top             =   8070
            Visible         =   0   'False
            Width           =   12495
            _Version        =   393216
            _ExtentX        =   22040
            _ExtentY        =   2672
            _StockProps     =   64
            BackColorStyle  =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   4
            MaxRows         =   14
            OperationMode   =   2
            ScrollBars      =   2
            SelectBlockOptions=   0
            ShadowColor     =   13697023
            SpreadDesigner  =   "frmInterface.frx":F7FA
         End
         Begin MSComCtl2.DTPicker dtpWorkDt 
            Height          =   315
            Left            =   18180
            TabIndex        =   59
            Top             =   330
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   66191361
            CurrentDate     =   40248
         End
         Begin MSComCtl2.DTPicker dtpToday 
            Height          =   315
            Left            =   13650
            TabIndex        =   111
            Top             =   270
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����ü"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   66191360
            CurrentDate     =   40457
         End
         Begin VB.Label Label1 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�˻�����"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
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
            Left            =   12720
            TabIndex        =   112
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Label3 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5070
            TabIndex        =   70
            Top             =   360
            Width           =   165
         End
         Begin VB.Label Label13 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "W/N"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   4020
            TabIndex        =   41
            Top             =   360
            Width           =   420
         End
         Begin VB.Label Label12 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2400
            TabIndex        =   34
            Top             =   330
            Width           =   105
         End
         Begin VB.Label Label1 
            Appearance      =   0  '���
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '����
            Caption         =   "�۾�����"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   32
            Top             =   330
            Width           =   780
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  '�� ����
      BackColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   16695
      TabIndex        =   23
      Top             =   0
      Width           =   16755
      Begin VB.CommandButton cmdIFTrans 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16830
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   9405
         TabIndex        =   55
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optSaveResult 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����ü"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   10185
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label5 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   8340
         TabIndex        =   56
         Top             =   90
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '���
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '����
         Caption         =   " KRYPTOR Interface"
         BeginProperty Font 
            Name            =   "����"
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
         TabIndex        =   27
         Top             =   90
         Width           =   1965
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12180
         Picture         =   "frmInterface.frx":FCA9
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13380
         Picture         =   "frmInterface.frx":10233
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14670
         Picture         =   "frmInterface.frx":107BD
         Top             =   90
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Port"
         Height          =   195
         Index           =   0
         Left            =   11640
         TabIndex        =   26
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Send"
         Height          =   195
         Left            =   12765
         TabIndex        =   25
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Receive"
         Height          =   195
         Left            =   13800
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '�Ʒ� ����
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   10755
      Width           =   16755
      _ExtentX        =   29554
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   10769
            MinWidth        =   10761
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
            TextSave        =   "2016-06-30"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            TextSave        =   "���� 3:39"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnMain 
      Caption         =   "����"
      Begin VB.Menu MnExit 
         Caption         =   "����"
      End
   End
   Begin VB.Menu MnConfig 
      Caption         =   "����"
      Begin VB.Menu MnTConfig 
         Caption         =   "��ż���"
      End
      Begin VB.Menu MnExamConfig 
         Caption         =   "�ڵ弳��"
      End
   End
   Begin VB.Menu MnTrans 
      Caption         =   "�����ϱ���"
      Begin VB.Menu MnTransAuto 
         Caption         =   "�ڵ�������"
      End
      Begin VB.Menu MnTransManual 
         Caption         =   "����������"
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
Const colSpecNo = 0 '�̻��
Const colCheckBox = 1
Const colSeqNo = 2
Const colOrdDate = 3
Const colWN = 4
Const colWK = 5
Const colBarcode = 6
Const colRack = 7
Const colPos = 8
Const colSale = 9
Const colCST = 10
Const colSPC = 11
Const colPID = 12
Const colPName = 13
Const colSex = 14
Const colAge = 15
Const colOCnt = 16
Const colRCnt = 17
Const colState = 18

'Const colrRslt = 19
'Const colrHL = 20
'Const colrOldResult = 21
'Const colrOldBar = 22
'Const colrLow = 23
'Const colrHigh = 24

Const colrRslt = 19
Const colrDil = 20
Const colrQC = 21
Const colrHL = 22
Const colrOldResult = 23
Const colrOldBar = 24
Const colrLow = 25
Const colrHigh = 26




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

Const colHL = 7
Const colPanic = 8
Const colDelta = 9
Const colOldRslt = 10
Const colOldBarcode = 11
Const colLowLimit = 12
Const colHighLimit = 13
Const colFLAG = 14

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
Dim blnBatch        As Boolean
'===============================

Dim strPrevRegNo    As String
Dim blnSame As Boolean

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

Private Sub cmdAddRow_Click()

    vasWorkList.MaxRows = vasWorkList.MaxRows + 1
    vasWorkList.RowHeight(-1) = 13
    
End Sub

'Private Sub cmdDownload_Click()
'    Dim intRow As Integer
'    Dim j  As Integer
'
'    j = 0
'    With vasWorkList
'        For intRow = 1 To .DataRowCnt
'            .Row = intRow
'            .Col = colCheckBox
'            If .Value = 1 Then
'                vasID.MaxRows = vasID.MaxRows + 1
'
'                .Col = colBarcode
'                SetText vasID, txtNum, vasID.MaxRows, colSeqNo
'                SetText vasID, Trim(.Text), vasID.MaxRows, colBarcode
'                Call GetSampleInfoW(vasID.MaxRows)                                '5,6,7,8
'
'                'Call .DeleteRows(intRow, intRow)
'                '.MaxRows = .MaxRows - 1
'                '.Action = ActionDeleteRow
''                .MaxRows = .MaxRows - 1
'
'                txtNum = txtNum + 1
'
'                .Col = colCheckBox
'                .Value = "0"
'
'            End If
'        Next
'    End With
'
'
'End Sub

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
    Dim blnWrite As Variant
    
    ClearSpread vasPrint

    blnWrite = False
    For iRow = 1 To vasRID.DataRowCnt
        vasRID.Row = iRow
        vasRID.Col = 1

        If vasRID.Value = 1 Then
            If blnWrite = False Then
                For j = 1 To 25
                    SetText vasPrint, Trim(GetText(vasRID, 0, j)), 0, j
                Next
            End If
            
            For j = 1 To 25
                SetText vasPrint, Trim(GetText(vasRID, iRow, j)), iRow, j
            Next
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "������ �ڷᰡ �����ϴ�.", , "�� ��"
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

' Excel Object Library �� �����մϴ�.
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
   ' txtNum = 0
    
    SetForeColor vasWorkList, 1, vasWorkList.MaxRows, 1, vasWorkList.MaxCols, 0, 0, 0
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    vasID.MaxRows = 0
    vasRes.MaxRows = 0
    vasWorkList.MaxRows = 0
'    dtptoday = Format(CDate(Date), "yyyy/mm/dd")
    
    txtCass = "A"
    txtPos = "1"
    
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

Private Sub cmdOrder_Click()
    Dim intRow As Integer
    
    With vasWorkList
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = colCheckBox
            If .Value = "1" Then
                strState = "Q"
                intSndPhase = 1
                intFrameNo = 1
                blnBatch = True
                comEqp.Output = ENQ
                Exit Sub
            End If
        Next
    End With
    
    MsgBox "������ ��ü�� �����Ͻð� ���������� �ϼ���", vbOKOnly + vbCritical, Me.Caption
    
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

Private Sub cmdPrint_Click()

    vasRID.PrintOrientation = PrintOrientationLandscape '�������
    vasRID.Action = 13
    
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
    dtpExamDateTo = Date
    
End Sub

Private Sub cmdRSch_Click()
    Dim iRow As Long

    ClearSpread vasRID
    ClearSpread vasRRes
    
    vasRID.MaxRows = 0
    
    Call chkRAll_Click
    
    'SELECT ó�� '' �� üũ�ڽ�
'         SQL = " SELECT DISTINCT '','',EXAMDATE,WORKNO,WORKKEY,BARCODE,DISKNO, POSNO, SALETEAM,DEALTEAM,SAMPLETYPE,PID, PNAME, PSEX, PAGE,'','',SENDFLAG,RESULT,REFFLAG,DELTAVALUE,REFVALUE,OLDRESULT,OLDBARCODE,REFLOW,REFHIGH " & vbCrLf
          SQL = " SELECT DISTINCT '','',EXAMDATE,WORKNO,WORKKEY,BARCODE,DISKNO, POSNO, SALETEAM,DEALTEAM,SAMPLETYPE,PID, PNAME, PSEX, PAGE,SENDDATE,'',SENDFLAG,RESULT,REFFLAG,DELTAVALUE,REFVALUE,OLDRESULT,OLDBARCODE,REFLOW,REFHIGH " & vbCrLf
    SQL = SQL & "   FROM PATRESULT " & vbCrLf
    SQL = SQL & "  WHERE EXAMDATE BETWEEN '" & Format(dtpExamDate, "YYYYMMDD") & "' AND '" & Format(dtpExamDateTo, "YYYYMMDD") & "' " & vbLf
    SQL = SQL & "    AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    If chkSave(0).Value = "1" Then
        SQL = SQL & "    AND SENDFLAG IN ('0','1','2') " & vbCrLf
    Else
        SQL = SQL & "    AND SENDFLAG IN ('0','1') " & vbCrLf
    End If
    
    If chkQC.Value = "1" Then
        SQL = SQL & "   AND BARCODE LIKE '%PCT%' "
    Else
        SQL = SQL & "   AND NOT BARCODE LIKE '%PCT%' "
    End If
    SQL = SQL & " ORDER BY EXAMDATE,SENDDATE "
'    SQL = SQL & "  GROUP BY EXAMDATE, BARCODE, DISKNO, POSNO, PID, PNAME, PSEX, PAGE, SENDFLAG "
          
    Res = GetDBSelectVas(gLocal, SQL, vasRID)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For iRow = 1 To vasRID.DataRowCnt
        Select Case Trim(GetText(vasRID, iRow, colState))
            Case "0": SetText vasRID, "Result", iRow, colState
            Case "1": SetText vasRID, "���", iRow, colState
            Case "2": SetText vasRID, "Trans", iRow, colState
                      SetBackColor vasRID, iRow, iRow, 1, colState, 202, 255, 112
        End Select
        
        If Trim(GetText(vasRID, iRow, colrHL)) = "L" Then
            SetBackColor vasRID, iRow, iRow, colrHL, colrHL, 0, 0, 255
        End If
        If Trim(GetText(vasRID, iRow, colrHL)) = "H" Then
            SetBackColor vasRID, iRow, iRow, colrHL, colrHL, 255, 0, 0
        End If
        
        If Trim(GetText(vasRID, iRow, colOCnt)) <> "" Then
            SetBackColor vasRID, iRow, iRow, colPName, colPName, 202, 255, 112
        End If
                
    Next iRow
    
    vasRID.RowHeight(-1) = 13
    
End Sub

Private Sub cmdRTrans_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasRID.DataRowCnt
        vasRID.Row = lRow
        vasRID.Col = 1
        If vasRID.Value = 1 Then
            Res = SaveTransDataR(lRow)
            'Res = SaveTransDataR(lRow)
        
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
                      " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
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
    Dim strDob As String
    Dim strSex As String
    Dim strIDNO As String
    Dim strOldRegNo As String
    Dim strOldResult As String
    Dim strOldDate As String
    Dim strBarNo As String
    Dim RecordsAffected As Long
    
    'intRow = vasWorkList.MaxRows
    vasWorkList.MaxRows = 0
    txtCass = "A"
    txtPos = "1"
    
    '-- ���� �˻��ڵ� ã��
'    Debug.Print gAllExam
    
    '-- �˻����� ��������
    SQL = ""
'    SQL = SQL & " SELECT REQNO,ITEMCD,SAMPCD,SAMPNM,WRKKEY,LABEMP,WRKDTE,WORKNO,REQDTE,DEPTCD,ITEMNM,PATNM,HOSNO,IDNO,BRCCD,BRCNM,CSTCD,CSTNM,LABRES,ETC2,DIALYSISYN,URIVOL,MNGNO,PACKCD1 " & vbLf
    SQL = SQL & " SELECT * " & vbLf
    SQL = SQL & "   FROM MCHORDER " & vbLf
    SQL = SQL & "  WHERE WRKDTE BETWEEN '" & pFrDt & "' AND '" & pToDt & "' " & vbLf
    SQL = SQL & "    AND WORKNO BETWEEN '" & Trim(txtStartNum.Text) & "' AND '" & Trim(txtStopNum.Text) & "' " & vbLf
    SQL = SQL & "    AND LABEMP = '" & gUserID & "' " & vbLf
    SQL = SQL & "    AND ITEMCD IN (" & gAllExam & ") " & vbLf
    If chkSave(1).Value = "0" Then
        SQL = SQL & "    AND (LABRES = '' OR LABRES IS NULL)"
    End If
    SQL = SQL & "  ORDER BY WRKDTE, WORKNO "

    '-- Record Count ������
    cn_Ser.CursorLocation = adUseClient
    Set RS = cn_Ser.Execute(SQL, , 1)
    
    FraSearch.Visible = True
    Screen.MousePointer = 11
    If RS.RecordCount > 0 Then
        progBar.Max = RS.RecordCount
    End If
    DoEvents
    
    Do Until RS.EOF
        progBar.Value = i
        i = i + 1
        intRow = intRow + 1
        vasWorkList.MaxRows = intRow
        
        If Trim(RS.Fields("LABRES").Value) & "" <> "" Then
            SetForeColor vasWorkList, intRow, intRow, 1, colrRslt, 255, 100, 0
        End If
        
        SetText vasWorkList, "0", intRow, colCheckBox
        SetText vasWorkList, Trim(RS.Fields("WORKNO").Value) & "", intRow, colWN
        SetText vasWorkList, Trim(RS.Fields("WRKKEY").Value) & "", intRow, colWK
        SetText vasWorkList, Trim(RS.Fields("WRKDTE").Value) & "", intRow, colOrdDate
        strBarNo = Trim(RS.Fields("REQNO").Value) & ""
        SetText vasWorkList, strBarNo, intRow, colBarcode
'        SetText vasWorkList, Trim(RS.Fields("BRCNM").Value) & "", intRow, colSale
        SetText vasWorkList, Trim(RS.Fields("BRCCD").Value) & "", intRow, colSale
        SetText vasWorkList, Trim(RS.Fields("CSTNM").Value) & "", intRow, colCST
        SetText vasWorkList, Trim(RS.Fields("SAMPNM").Value) & "", intRow, colSPC
        SetText vasWorkList, Trim(RS.Fields("HOSNO").Value) & "", intRow, colPID
        SetText vasWorkList, Trim(RS.Fields("PATNM").Value) & "", intRow, colPName
        '-- ��� ������ ������
        SetText vasWorkList, Trim(RS.Fields("LABRES").Value) & "", intRow, colrRslt
        strIDNO = Trim(RS.Fields("IDNO").Value) & ""
        Select Case Mid(strIDNO, 7, 1)
            Case "1"
                strSex = "M"
                strDob = "19" & Mid(strIDNO, 1, 6)
            Case "2"
                strSex = "F"
                strDob = "19" & Mid(strIDNO, 1, 6)
            Case "3"
                strSex = "M"
                strDob = "20" & Mid(strIDNO, 1, 6)
            Case "4"
                strSex = "F"
                strDob = "20" & Mid(strIDNO, 1, 6)
            Case "5"
                strSex = "M"
                strDob = "19" & Mid(strIDNO, 1, 6)
            Case "6"
                strSex = "F"
                strDob = "19" & Mid(strIDNO, 1, 6)
            Case "7"
                strSex = "M"
                strDob = "19" & Mid(strIDNO, 1, 6)
            Case "8"
                strSex = "F"
                strDob = "19" & Mid(strIDNO, 1, 6)
            Case "9"
                strSex = "M"
                strDob = "18" & Mid(strIDNO, 1, 6)
            Case "0"
                strSex = "F"
                strDob = "18" & Mid(strIDNO, 1, 6)
            Case Else
                strSex = ""
                strDob = ""
        End Select
        SetText vasWorkList, strSex, intRow, colSex
        SetText vasWorkList, strDob, intRow, colAge
        
        SetText vasWorkList, txtCass, intRow, colRack
        SetText vasWorkList, txtPos, intRow, colPos

        txtPos = txtPos + 1
        If txtPos = "17" Then
            Select Case txtCass
            Case "A": txtCass = "B"
            Case "B": txtCass = "E"
            Case "E": txtCass = "F"
            Case "F": txtCass = ""
            Case Else: txtCass = ""
            End Select
            txtPos = "1"
        End If
                
                '-- ������� ��ȸ
              SQL = " Select PASTREQNO, LABRES,INPDTE "
        SQL = SQL & "   From PastRes "
        SQL = SQL & "  Where REQNO = '" & strBarNo & "'"
        SQL = SQL & "    And ITEMCD = 'X274'"
        SQL = SQL & "  Order by PASTREQNO DESC"
        Res = GetDBSelectColumn(gServer, SQL)
        
        If Res > 0 Then
            strOldRegNo = Trim(gReadBuf(0))     '���� �Ƿڹ�ȣ
            strOldResult = Trim(gReadBuf(1))    '���� ���
            strOldDate = Trim(gReadBuf(0))     '���� �����
        End If

        SetText vasWorkList, strOldResult, intRow, colrOldResult
        If strOldRegNo <> "" Then
            SetText vasWorkList, strOldRegNo & "-" & strOldDate, intRow, colrOldBar
        End If
        RS.MoveNext
    Loop
    
    FraSearch.Visible = False
    Screen.MousePointer = 0

    vasWorkList.RowHeight(-1) = 13
    vasWorkList.Row = 1
    
    RS.Close
    Set RS = Nothing
    
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

Private Sub cmdSave_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasWorkList.DataRowCnt
        vasWorkList.Row = lRow
        vasWorkList.Col = 1
        If vasWorkList.Value = 1 Then
            Res = SaveTransDataR(lRow)
        
            If Res = -1 Then
                SetForeColor vasWorkList, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasWorkList, "Failed", lRow, colState
            ElseIf Res = 0 Then
            
            Else
                vasWorkList.Row = lRow
                vasWorkList.Col = 1
                vasWorkList.Value = 1
                
                SetBackColor vasWorkList, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasWorkList, "Trans", lRow, colState
                
                SQL = " UPDATE PATRESULT SET " & vbCrLf & _
                      " SENDFLAG = '2' " & vbCrLf & _
                      " WHERE EQUIPNO = '" & gEquipCode & "' " & vbCrLf & _
                      "   AND BARCODE = '" & Trim(GetText(vasWorkList, lRow, colBarcode)) & "' "
                Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
            vasWorkList.Row = lRow
            vasWorkList.Col = 1
            vasWorkList.Value = 0
        End If
    Next lRow
End Sub

Private Sub cmdSearch_Click()
                
    If dtpStartDt > dtpStopDt Then
        MsgBox "��ȸ�Ⱓ �����Դϴ�.", vbOKOnly + vbCritical, Me.Caption
        dtpStartDt.SetFocus
        Exit Sub
    End If
    
    If txtStartNum > txtStopNum Then
        MsgBox "W/N�Է� �����Դϴ�.", vbOKOnly + vbCritical, Me.Caption
        txtStartNum.SetFocus
        Exit Sub
    End If
    
    Call GetWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    
    vasID.RowHeight(-1) = 13
    txtStartNum.SetFocus

End Sub




Private Sub imgPort_DblClick()
    
'    '-- ���߽ÿ��� Remark Ǯ� �׽�Ʈ����
'    If FrmHideControl.Visible = True Then
'        Me.Width = 15435
'        FrmHideControl.Visible = False
'    Else
'        Me.Width = 25000
'        FrmHideControl.Visible = True
'    End If

End Sub



Private Sub Label1_DblClick(Index As Integer)

    If Index = 2 Then
        If FraLog.Visible = True Then
            FraLog.Visible = False
        Else
            FraLog.Visible = True
        End If
    End If
    
End Sub

Private Sub lblclear_Click()
    lblChangePID.Caption = ""
    lblChangeBar.Caption = ""
'    lblBarcode(0).Caption = ""
    lblBarcode(1).Caption = ""
'    lblPname(0).Caption = ""
    lblPname(1).Caption = ""
End Sub

Private Sub Command16_Click()
    Dim i As Long
    Dim lsChar As String
        
'    strBuffer = ""
'strBuffer = strBuffer & "1H|\^&|||KRYPTOR^AUTOMATE KRYPTOR|||||LIS||P|1|199709011638"
'strBuffer = strBuffer & "20"
'2P|1|||||||U
'6C
'3O|1|02315000^01^10||^^^CEA^^1|R||||||A||||||||01^10||||||F
'8A
'4R|1|^^^CEA^^1^^F|126.854|||H||F|||19970901151020|19970901163000|517017000187
'E7
'5C|1|I|40
'2B
'6L|1|F
'05
'

    Call comEqp_OnComm
        

End Sub



Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
On Error GoTo Rst

    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
        
    'Me.Height = 11520
    'Me.Width = 15435
    
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
        frmInterface.StatusBar1.Panels(2).Text = "���� �Ǿ����ϴ�"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        lblStatus = "�۾���.."
    Else
        frmInterface.StatusBar1.Panels(2).Text = "���� ���� �ʾҽ��ϴ�"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
'        lblStatus = "�۾� �����.."
    End If

    If Not Connect_Local Then
        MsgBox "������� �ʾҽ��ϴ�."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    '-- osw �߰�
'    For i = 1 To 1
'        If Not Connect_PRServer Then
'            MsgBox "������� �ʾҽ��ϴ�."
'            cn_Server_Flag = False
'            Exit Sub
'        Else
'            cn_Server_Flag = True
'        End If
'    Next
    
    
    GetExamCode
    
    dtpToday = Date
    dtpStartDt = Date
    dtpStopDt = Date
    dtpWorkDt = Date
    
    txtStartNum.Text = "0000"
    txtStopNum.Text = "9999"
    
    sDate = Format(DateAdd("y", CDate(dtpToday.Value), -60), "yyyy-mm-dd")
    
          SQL = "DELETE FROM PATRESULT "
    SQL = SQL & " WHERE examdate < " & sDate
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
    blnBatch = False
    Timer1.Enabled = False
    '==============================
    
    frmInterface.StatusBar1.Panels(2).Text = Winsock1.LocalIP
    
    Exit Sub
    
Rst:
    If Err.Number = "8002" Then
        If (MsgBox("��Ʈ ��ȣ�� �߸��Ǿ����ϴ�." & vbNewLine & vbNewLine & "   ��� �����Ͻðڽ��ϱ�?", vbYesNo + vbCritical, Me.Caption)) = vbYes Then
            Resume Next
        Else
            End
        End If
    End If
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

'    Call dce_close_env      ' Server�� ������ ���� ��
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
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim intRow As Integer
    Dim strOutput As String     '�۽��� ������
    Dim varGettxt As Variant
    Dim strPID As String
    Dim strDob As String
    Dim strSex As String
    Dim strBarNo As String
    Dim strOldResult As String
    Dim strOldBar As String
    Dim strOldDate As String
    Dim strState  As String
    Dim strPrevResult As String
    
'<STX>1H|\^&|||LIS|||||KRYPTOR||P|1|199709011040<CR ><ETX>xx<CR ><LF >
'<STX>2P|1|9401134001|||DUPONT^MartheI||19580613|F<CR ><ETX>xx<CR ><LF >
'<STX>3O|1|01500||^^^AFP\^^^FSH|R||||||A||||||||||||||Q<CR ><ETX>xx<CR ><LF >
'<STX>4L|1|F<CR ><ETX>xx<CR ><LF >
    
    With vasWorkList
        intRow = mOrder.Seq
        Select Case intSndPhase
            Case 1  '## Header
                strOutput = intFrameNo & "H|\^&|||LIS|||||KRYPTOR||P|1|" & Format(Now, "yyyymmddhhmm") & vbCr & ETX
                intSndPhase = 2
                intFrameNo = intFrameNo + 1
            Case 2  '## Patient
                Call .GetText(colPID, intRow, varGettxt): strPID = CStr(varGettxt)
                Call .GetText(colAge, intRow, varGettxt): strDob = CStr(varGettxt)
                Call .GetText(colSex, intRow, varGettxt): strSex = CStr(varGettxt)
                
'                strPID = ""
'                strDob = ""
'                strSex = "M"

                strOutput = intFrameNo & "P|1|" & strPID & "|||I|" & strDob & "|" & strSex & vbCr & ETX
                intSndPhase = 3
                intFrameNo = intFrameNo + 1
            Case 3  '## Order
'                 MSComm1.Output = STX & "3O|1|549116451352^01^10  ||^^^CEA^^1     |R||||||A||||||||||||||F" & vbCr & ETX & "94" & vbCr & vbLf
                '-- Auto Dilution
                strOutput = intFrameNo & "O|1|" & mOrder.BarNo & "||^^^PCT^^1" & "|R||||||A||||||||||||||Q" & vbCr & ETX
                
                '-- Query Mode
                'O|1|01201||^^^CEA^^1\^^^AFP^^2^19970512V26.1\^^^LH|R||||||A||||||||||||||Q<CR >
                '-- Batch Mode
                'O|1|01201^01^10||^^^CEA^^1\^^^AFP^^2^19970512V26.1\^^^LH|R||||||A||||||||||||||O<CR >

                intSndPhase = 4
                intFrameNo = intFrameNo + 1
            Case 4  '## Termianator
                strOutput = intFrameNo & "L|1|F" & vbCr & ETX
                intSndPhase = 5
                intFrameNo = intFrameNo + 1
            Case 5  '## EOT
                strState = ""
                comEqp.Output = EOT
                SetRawData "[Tx]" & EOT
                intFrameNo = 1
                Exit Sub
        End Select
    End With
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    comEqp.Output = strOutput
    Debug.Print strOutput
    SetRawData "[Tx]" & strOutput
    
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder_Batch()
    Dim intRow As Integer
    Dim strOutput As String     '�۽��� ������
    Dim varGettxt As Variant
    Dim strPID As String
    Dim strDob As String
    Dim strSex As String
    Dim strBarNo As String
    Dim strOldResult As String
    Dim strOldBar As String
    Dim strOldDate As String
    Dim strState  As String
    Dim strPrevResult As String
    Dim strCass As String
    Dim strPos  As String
    
'<ENQ>
'<ACK>
'<STX>1H|\^&|||LIS|||||KRYPTOR||P|1|199709011100<CR ><ETX>xx<CR ><LF >
'<ACK>
'<STX>2P|1|9401134002|||RIEUSSET^HENRI||19260430|M<CR ><ETX>xx<CR ><LF >
'<ACK>
'<STX>3C|1|P|mon commentaire patient limite a 70<CR ><ETX>xx<CR ><LF >
'<NAK>
'<STX>3C|1|P|mon commentaire patient limite a 70<CR ><ETX>xx<CR ><LF >
'<ACK>
'<STX>4O|1|02315000^01^10||^^^CEA^^1^19970512V46.2\^^^FSH^^^19970104V450|R||||||A||||||||||||||O<CR >
'<ETX>xx<CR ><LF >
'<ACK>
'<STX>5L|1|F<CR ><ETX>xx<CR ><LF >
'<ACK>
'<EOT>


    With vasWorkList
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = colState: strState = Trim(.Value)
            .Col = colCheckBox
            If .Value = "1" And strState <> "����" Then
                Select Case intSndPhase
                    Case 1  '## Header
                        strOutput = intFrameNo & "H|\^&|||LIS|||||KRYPTOR||P|1|" & Format(Now, "yyyymmddhhmm") & vbCr & ETX
                        intSndPhase = 2
                        intFrameNo = intFrameNo + 1
                        Exit For
                    Case 2  '## Patient
                        Call .GetText(colRack, intRow, varGettxt): strCass = varGettxt
                        Call .GetText(colPos, intRow, varGettxt): strPos = varGettxt
                    
                        Call .GetText(colPID, intRow, varGettxt): strPID = CStr(varGettxt)
                        Call .GetText(colAge, intRow, varGettxt): strDob = CStr(varGettxt)
                        Call .GetText(colSex, intRow, varGettxt): strSex = CStr(varGettxt)
                        strOutput = intFrameNo & "P|1|" & strCass & strPos & "|||^I||" & strDob & "|" & strSex & vbCr & ETX
                        intSndPhase = 4
                        intFrameNo = intFrameNo + 1
                        Exit For
'                    Case 3  '## Comment
'                        '<STX>3C|1|P|mon commentaire patient limite a 70<CR ><ETX>xx<CR ><LF >
'                        strOutput = intFrameNo & "C|1|" & vbCr & ETX
'                        intSndPhase = 4
'                        intFrameNo = intFrameNo + 1
'                        Exit For
                    Case 4  '## Order
                        Call .GetText(colRack, intRow, varGettxt): strCass = varGettxt
                        Call .GetText(colPos, intRow, varGettxt): strPos = varGettxt

                        Call .GetText(colBarcode, intRow, varGettxt): strBarNo = CStr(varGettxt)
                        Call .GetText(colrOldResult, intRow, varGettxt): strOldResult = CStr(varGettxt)
                        Call .GetText(colrOldBar, intRow, varGettxt)
                        If CStr(varGettxt) <> "" Then
                            strOldBar = mGetP(CStr(varGettxt), 1, "-")
                            strOldDate = mGetP(CStr(varGettxt), 2, "-")
                        End If
                        'Bar to REQ
                        'SQL = "SELECT TO_CHAR(TO_DATE('20000101','YYYYMMDD') + TO_NUMBER(SUBSTR('" & strBarNo & "', 1, 4)), 'YYYYMMDD') || SUBSTR('" & strBarNo & "', 5, 7) as V_REQNO FROM DUAL"
                        
                        'REQ to Bar
                        SQL = "SELECT LPAD(TO_DATE(SUBSTR('" & strBarNo & "', 1, 8),'YYYYMMDD') - TO_DATE('20000101','YYYYMMDD'), 4, '0') || SUBSTR('" & strBarNo & "', 9, 7) as V_BARCODE FROM DUAL"
                        Res = GetDBSelectColumn(gServer, SQL)
                        If Res > 0 Then
                            strBarNo = Trim(gReadBuf(0))     '�Ƿڹ�ȣ
                        End If
                        If strOldDate <> "" And strOldResult <> "" Then
                            strPrevResult = strOldDate & "V" & strOldResult
                        End If
                        '-- Auto Dilution
                        strOutput = intFrameNo & "O|1|" & strBarNo & "^" & strCass & "^" & strPos & "||^^^PCT^^1^|R||||||A||||||||||||||O" & vbCr & ETX
                        'strOutput = intFrameNo & "O|1|" & strBarNo & "^01^10||^^^PCT^^1^|R||||||A||||||||||||||O" & vbCr & ETX
                        
                        'strOutput = intFrameNo & "O|1|" & strBarNo & "||^^^PCT^^1^|R||||||A||||||||||||||O" & vbCr & ETX
                        
                        '-- ���� strOutput = intFrameNo & "O|1|" & strBarNo & "||^^^PCT^^1" & "|R||||||A||||||||||||||O" & vbCr & ETX


                        'strOutput = intFrameNo & "O|1|" & strBarNo & "^" & strCass & "^" & strPos & "||^^^CEA^^1^19970512V46.2\^^^PCT^^^" & "|R||||||A||||||||||||||O" & vbCr & ETX
                        'strOutput = intFrameNo & "O|1|" & strBarNo & "^01^10||^^^PCT^^1^" & strPrevResult & "|R||||||A||||||||||||||O" & vbCr & ETX
                        
                        '-- Query Mode
                        'O|1|01201||^^^CEA^^1\^^^AFP^^2^19970512V26.1\^^^LH|R||||||A||||||||||||||Q<CR >
                        '-- Batch Mode
                        'O|1|01201^01^10||^^^CEA^^1\^^^AFP^^2^19970512V26.1\^^^LH|R||||||A||||||||||||||O<CR >
'                        Field 9.3 : 1th component : sample identification with a maximum of 15 alpha-numeric
'                        characters.
'                        2nd component : cassette number - numeric , from 01 to 99
'                        3rd component : position number - numeric from 01 to 10
'                        Those 2 components are optional. Their value are transmitted to KRYPTOR if
'                        specified.
'                        Field 9.5 : 4th component : mnemonic KRYPTOR test code .
'                        (see Analyte codes)
'                        6th component : dilution . When no dilution is specified, 1 is the default value
'                        7th component : previous result with the following format : AAAAMMJJV999.99
'                        Date AAAAMMJJ
'                        letter �� V �� followed by value
'                        The previous result value is transmitted to Kryptor
'                        Field 9.6 : �� R �� Routine
'                        �� A �� Very urgent
'                        �� S �� Urgent
'                        both A and S are handled by Kryptor as Stat samples
'                        Field 9.12 : �� A �� Add the test
'                        �� Q �� Consider that sample as a quality control sample. (not implemented in
'                        KIM V 1.20)
'                        In this case, both name and PID will be replaced by text �� Control ��.
'                        Field 9.26 : �� O �� For a "batch" mode request.
                        
                        intSndPhase = 5
                        intFrameNo = intFrameNo + 1
                        Exit For
                    Case 5  '## Termianator
                        strOutput = intFrameNo & "L|1|F" & vbCr & ETX
                        intSndPhase = 6
                        intFrameNo = intFrameNo + 1
                        Exit For
                        
                    Case 6  '## EOT
                        strState = ""
                        comEqp.Output = EOT
                        SetRawData "[Tx]" & EOT
                        intFrameNo = 1
                        .Row = intRow
                        .Col = colCheckBox: .Value = "0"
                        .Col = colState: .Value = "����"
                        Timer1.Interval = 500
                        Timer1.Enabled = True
                        Exit Sub
                End Select
            End If
        Next
        
        strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
        comEqp.Output = strOutput
        Debug.Print strOutput
        SetRawData "[Tx]" & strOutput

    End With
    
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڿ��� CheckSum�� ����
'   �μ� :
'       - pMsg : ���ڿ�
'   ��ȯ : CheckSum
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

'-- ���ݳ�¥�� �˻����� ���Ѵ�
Function DateCompare(ByVal FDate As String) As String
    
    DateCompare = FDate
    If FDate <> Format(Now, "yyyymmdd") Then
        DateCompare = Format(Now, "yyyymmdd")
    End If
    
End Function

Private Sub Timer1_Timer()
    Dim intRow As Integer
    
    With vasWorkList
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = colCheckBox
            If .Value = "1" Then
                strState = "Q"
                intSndPhase = 1
                intFrameNo = 1
                blnBatch = True
                comEqp.Output = ENQ
                Exit For
            End If
        Next
    End With
    
    Timer1.Enabled = False

End Sub

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
            
            txtData = txtData & Buffer
            
            txtComLog = txtComLog & Buffer
            
            SetRawData "[Rx]" & Buffer
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
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
                            Case ACK
                                '-- ��񿡼� �Ѿ�� �ð��� �쿬�� 11:59:59�ʳ� ���Ͽ� ����� �ð��� ���
                                '-- ��� ����� �������� ������ �� �����Ƿ� ��¥�� �ǽð� ������Ʈ �Ѵ�.
                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
                                dtpToday.Value = Format(strDate, "####-##-##")
                                
                                DoEvents
                                
                                If strState = "Q" Then
                                    If blnBatch = True Then
                                        Call SendOrder_Batch
                                    Else
                                        Call SendOrder
                                    End If
                                End If
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                Erase strRecvData
                                comEqp.Output = ACK
                                SetRawData "[Tx]" & ACK
                            Case STX
                            '    intBufCnt = 1
                            '    Erase strRecvData
                             '   ReDim Preserve strRecvData(intBufCnt)
                            Case ETB
                                blnIsETB = True
                                intPhase = 3
                            Case ETX
                             '   intBufCnt = intBufCnt + 1
                             '   ReDim Preserve strRecvData(intBufCnt)
                                intPhase = 3
                            Case vbCr
                                intBufCnt = intBufCnt + 1
                                ReDim Preserve strRecvData(intBufCnt)
                            Case vbLf
'                            Case EOT
'                                strDate = DateCompare(Format(CDate(dtpToday.Value), "yyyymmdd"))
'                                dtpToday.Value = Format(strDate, "####-##-##")
'
'                                DoEvents
'
'                                Call EditRcvDataASTM
'
'                                intPhase = 1
                            
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
'                                    Case vbLf
'                                        intPhase = 4
'                                        comEqp.Output = ACK
'                                        SetRawData "[Tx]" & ACK
                        End Select
                    Case 4      '## Termination Phase
                        Select Case BufChar
                            Case STX
                                intPhase = 2
                            Case EOT
                                '-- ��񿡼� �Ѿ�� �ð��� �쿬�� 11:59:59�ʳ� ���Ͽ� ����� �ð��� ���
                                '-- ��� ����� �������� ������ �� �����Ƿ� ��¥�� �ǽð� ������Ʈ �Ѵ�.
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
            EVMsg$ = "CTS ���� ����"
        Case comEvDSR
            EVMsg$ = "DSR ���� ����"
        Case comEvCD
            EVMsg$ = "CD ���� ����"
        Case comEvRing
            EVMsg$ = "��ȭ ���� �︮�� ��"
        Case comEvEOF
            EVMsg$ = "EOF ����"

        '���� �޽���
        Case comBreak
            ERMsg$ = "�ߴ� ��ȣ ����"
        Case comCDTO
            ERMsg$ = "�ݼ��� ���� �ð� �ʰ�"
        Case comCTSTO
            ERMsg$ = "CTS �ð� �ʰ�"
        Case comDCB
            ERMsg$ = "DCB �˻� ����"
        Case comDSRTO
            ERMsg$ = "DSR �ð� �ʰ�"
        Case comFrame
            ERMsg$ = "�����̹� ����"
        Case comOverrun
            ERMsg$ = "�и�Ƽ ����"
        Case comRxOver
            ERMsg$ = "���� ���� �ʰ�"
        Case comRxParity
            ERMsg$ = "�и�Ƽ ����"
        Case comTxFull
            ERMsg$ = "���� ���ۿ� ������ ����"
        Case Else
            ERMsg$ = "�� �� ���� ���� �Ǵ� �̺�Ʈ"
    End Select


End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��ȣ�� ���� �������� ��ȸ, ǥ��, �˻���������
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
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
            Exit For
        End If
    Next i
    
    If intRow < 0 Then
        intRow = vasWorkList.DataRowCnt + 1
        If vasWorkList.MaxRows < intRow Then
            vasWorkList.MaxRows = intRow
        End If
    End If
    
    mOrder.Seq = intRow
    '-- ���������� ǥ��
    Call SetText(vasWorkList, pBarNo, intRow, colBarcode)         '3  ���ڵ�
''    Call SetText(vasWorkList, mOrder.RackNo, intRow, colCST)     '4  Rack��ȣ
'    Call SetText(vasWorkList, mOrder.TubePos, intRow, colSPC)     '5  Pos��ȣ
    
    Call vasActiveCell(vasWorkList, intRow, colBarcode)
    Call ClearSpread(vasRes)
    
    '-- �˻��� ���� �������̺��� ������ ǥ��(for ��ũ����Ʈ)  '6,7,8,9
    If GetSampleInfoW(intRow) > 0 Then
        '-- ���ڵ��ȣ�� �����ϴ� �˻��ڵ� ��������(�μ� : ����ڵ�,���ڵ��ȣ)
        gOrderExam = GetOrderExamCode(gEquip, pBarNo)
        
        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetGetEquipExamCode_KRYPTOR(gEquip, pBarNo, intRow)
    End If
    

    '-- �˻�ä�η� ������ �����
    If Trim(strItems) = "" Then
        mOrder.NoOrder = True
        mOrder.Order = strItems
    Else
        mOrder.NoOrder = False
        mOrder.Order = strItems
    End If
    
    Call SetText(vasWorkList, "����", intRow, colState)         '12 �������

End Sub

'-----------------------------------------------------------------------------'
'   ��� :
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'-----------------------------------------------------------------------------'
Private Sub SetPatInfo(ByVal pBarNo As String, Optional ByVal intRow2 As Integer = 0)
    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    
    If intRow2 = 0 Then
        intRow = -1
        For i = 1 To vasWorkList.DataRowCnt
            If Trim(GetText(vasWorkList, i, colBarcode)) = pBarNo Then
                intRow = i
                '-- ���ڵ��ȣ�� �����ϴ� �˻��ڵ� ��������(�μ� : ����ڵ�,���ڵ��ȣ)
                gOrderExam = GetOrderExamCode(gEquip, pBarNo)
                Exit For
            End If
        Next i
    Else
        intRow = intRow2
        SetBackColor vasWorkList, intRow, intRow, colPName, colPName, 202, 255, 112
    End If
    
    If intRow < 0 Then
        intRow = vasWorkList.DataRowCnt + 1
        If vasWorkList.MaxRows < intRow Then
            vasWorkList.MaxRows = intRow
        End If
    End If

    '-- ���������� ǥ��
    Call SetText(vasWorkList, pBarNo, intRow, colBarcode)
    Call SetText(vasWorkList, mResult.RackNo, intRow, colRack)
    Call SetText(vasWorkList, mResult.TubePos, intRow, colPos)
    Call vasActiveCell(vasWorkList, intRow, colBarcode)

    '-- ����������� �����
    Call ClearSpread(vasRes)

    '-- �˻��� ���� �������̺��� ������ ǥ��(for ��ũ����Ʈ)  '6,7,8,9
    If GetSampleInfoW(intRow) > 0 Then
        '-- ���ڵ��ȣ�� �����ϴ� �˻��ڵ� ��������(�μ� : ����ڵ�,���ڵ��ȣ)
        gOrderExam = GetOrderExamCode(gEquip, pBarNo)
        
        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetGetEquipExamCode_CentaurCP(gEquip, pBarNo, intRow)
    End If
    
    '-- ���� Row
    gRow = intRow
    

End Sub


'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataASTM()
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ ���ڵ��ȣ
    Dim strSeq       As String   '������ Sequence
    Dim strRackNo    As String   '������ Rack Or Disk No
    Dim strTubePos   As String   '������ Tube Position
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���(����)
    Dim strIntResult As String   '������ ���(����)
    Dim strQCResult  As String   '������ ���(QC)
    Dim strFlag      As String   '������ Abnormal Flag
    Dim strComm      As String   '������ Comment
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
    Dim strOldRegNo As String
    Dim strOldResult As String
    Dim strReqNo As String
    Dim strDil   As String
    
    Dim strPrevStatus As String
    Dim strPrevResult As String
    Dim intRow2       As Integer
    
    'strRcvBuf = strRecvData(1)
    'varRcvBuf = Split(strRcvBuf, vbCr)
    
    For intCnt = 1 To UBound(strRecvData)
        strRcvBuf = strRecvData(intCnt)
        strType = Mid$(strRcvBuf, 2, 1)
        
        Select Case strType
            Case "Q"
                strState = "Q"
                '## ���ڵ��ȣ, SEQ, Disk No, Tube Position ��ȸ
                strTemp1 = mGetP(strRcvBuf, 3, "|")
                strBarNo = Trim$(mGetP(strTemp1, 2, "^"))
                '<STX>2Q|1|^01500|||ALL||||||||O<CR ><ETX>xx<CR ><LF >
                With mOrder
                    .BarNo = strBarNo
                    '.RackNo = mGetP(strTemp1, 1, "^")
                    '.TubePos = mGetP(strTemp1, 2, "^")
                    '.Seq = mGetP(strTemp1, 4, "^")
                End With
                
                If Len(strBarNo) >= 11 Then
                    'Bar to REQ
                    SQL = "SELECT TO_CHAR(TO_DATE('20000101','YYYYMMDD') + TO_NUMBER(SUBSTR('" & strBarNo & "', 1, 4)), 'YYYYMMDD') || SUBSTR('" & strBarNo & "', 5, 7) as V_REQNO FROM DUAL"
    
                    'REQ to Bar
                    'SQL = "SELECT LPAD(TO_DATE(SUBSTR('" & strBarNo & "', 1, 8),'YYYYMMDD') - TO_DATE('20000101','YYYYMMDD'), 4, '0') || SUBSTR('" & strBarNo & "', 9, 7) as V_BARCODE FROM DUAL"
                    Res = GetDBSelectColumn(gServer, SQL)
                    If Res > 0 Then
                        strBarNo = Trim(gReadBuf(0))     '�Ƿڹ�ȣ
                    End If
                End If
                
                If strBarNo = "" Then Exit Sub

                Call GetOrder(strBarNo)
                
            Case "O"
                strState = "O"
                strBarNo = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^"))
                strRackNo = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^"))
                strTubePos = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 3, "^"))
                
                If Len(strBarNo) >= 11 Then
                    'Bar to REQ
                    SQL = "SELECT TO_CHAR(TO_DATE('20000101','YYYYMMDD') + TO_NUMBER(SUBSTR('" & strBarNo & "', 1, 4)), 'YYYYMMDD') || SUBSTR('" & strBarNo & "', 5, 7) as V_REQNO FROM DUAL"
                    
                    'REQ to Bar
                    'SQL = "SELECT LPAD(TO_DATE(SUBSTR('" & strBarNo & "', 1, 8),'YYYYMMDD') - TO_DATE('20000101','YYYYMMDD'), 4, '0') || SUBSTR('" & strBarNo & "', 9, 7) as V_BARCODE FROM DUAL"
                    Res = GetDBSelectColumn(gServer, SQL)
                    If Res > 0 Then
                        strBarNo = Trim(gReadBuf(0))     '�Ƿڹ�ȣ
                    End If
                End If
                
                If strBarNo = "" Then Exit Sub
                
                intRow2 = 0
                blnSame = False
                For i = 1 To vasWorkList.DataRowCnt
                    vasWorkList.Row = i
                    vasWorkList.Col = colBarcode
                    If Trim(vasWorkList.Text) = strBarNo Then
                        SetText vasWorkList, strRackNo, i, colRack
                        SetText vasWorkList, strTubePos, i, colPos
                                                
                        '-- ����
                        strPrevStatus = GetText(vasWorkList, i, colState)
                        '-- ���
                        strPrevResult = GetText(vasWorkList, i, colrRslt)
                        
                        If (UCase(strPrevStatus) = "RESULT" Or UCase(strPrevStatus) = "TRANS") And strPrevResult <> "" Then
                            vasWorkList.MaxRows = vasWorkList.MaxRows + 1
                            gRow = vasWorkList.MaxRows
                            intRow2 = gRow
                            blnSame = True
                        Else
                            gRow = i
                        End If
                        Exit For
                    End If
                Next
                
                Call SetPatInfo(strBarNo, intRow2)
                
                With mResult
                    .BarNo = strBarNo
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                End With
                
                
                If gRow < 0 Then
                    Exit Sub
                End If
            
            Case "R"    '-- ���
                    
                '## ������ �˻��, ���, Abnormal Flag
                strIntBase = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"))
                strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                strDil = Trim(mGetP(mGetP(strRcvBuf, 3, "|"), 6, "^"))
                If strResult <> "" Then
                          SQL = ""
                    SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO,REFLOW,REFHIGH "
                    SQL = SQL & "  FROM EQPMASTER"
                    SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                    SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                    SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
                                            
                    Res = GetDBSelectColumn(gLocal, SQL)
                    
                    Dim strHighVal As String
                    Dim strLowVal  As String
                    Dim strLH      As String
                    
                    '-- ���� ���� ���
                    If Res > 0 Then
                        lsExamCode = Trim(gReadBuf(0))
                        lsExamName = Trim(gReadBuf(1))
                        lsSeqNo = Trim(gReadBuf(2))
                        strLowVal = Trim(gReadBuf(3))  'Low
                        strHighVal = Trim(gReadBuf(4)) 'High
                        
'                        lsResRow = vasRes.DataRowCnt + 1
'                        If vasRes.MaxRows < lsResRow Then
'                            vasRes.MaxRows = lsResRow
'                        End If
                        
                        '�Ҽ��� ó��, ��� ���� ó��
                        lsEquipRes = strResult
                        strResult = SetResult(strResult, strIntBase)
                        lsResult_Buff = strResult
                        
                        '����
                        strLH = ""
                        If IsNumeric(strResult) Then
                            If Val(strResult) < Val(strLowVal) Then
                                strLH = "L"
                                lsResult_Buff = "< " & strLowVal
                            End If
                            If Val(strResult) > Val(strHighVal) Then
                                strLH = "H"
                            End If
                        End If
                        
                        '-- Work List
                        SetText vasWorkList, "Result", gRow, colState                 '11 �������
                        
                        '-- ������� ��ȸ
                              SQL = " Select PASTREQNO, LABRES "
                        SQL = SQL & "   From PastRes "
                        SQL = SQL & "  Where REQNO = '" & strBarNo & "'"
                        SQL = SQL & "    And ITEMCD = '" & lsExamCode & "'"
                        SQL = SQL & "  Order by PASTREQNO DESC"
                        Res = GetDBSelectColumn(gServer, SQL)
                        
                        If Res > 0 Then
                            strOldRegNo = Trim(gReadBuf(0))     '���� �Ƿڹ�ȣ
                            strOldResult = Trim(gReadBuf(1))    '���� ���
                        End If
                        
                        '-- ��� List
                        SetText vasWorkList, lsResult_Buff, gRow, colrRslt          '���
                        SetText vasWorkList, strDil, gRow, colrDil          'Dilution
                        
                        SetText vasWorkList, strLH, gRow, colrHL               '���� High/Low
                        If strLH = "L" Then
                            SetBackColor vasWorkList, gRow, gRow, colrHL, colrHL, 0, 0, 255
                        End If
                        If strLH = "H" Then
                            SetBackColor vasWorkList, gRow, gRow, colrHL, colrHL, 255, 0, 0
                        End If
                        SetText vasWorkList, strOldResult, gRow, colrOldResult               '���� ���
                        SetText vasWorkList, strOldRegNo, gRow, colrOldBar               '���� �Ƿڹ�ȣ
                        SetText vasWorkList, strLowVal, gRow, colrLow               '����ġ ����
                        SetText vasWorkList, strHighVal, gRow, colrHigh               '����ġ ����
                        
                        '-- ���� ����
                        SetLocalDB gRow, gRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = "R"
                        
                    '-- ���� ���� ���
                    Else
                    
                              SQL = ""
                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO,REFLOW,REFHIGH "
                        SQL = SQL & "  FROM EQPMASTER"
                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
                                                                    
                        Res = GetDBSelectColumn(gLocal, SQL)
                        
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            strLowVal = Trim(gReadBuf(3))  'Low
                            strHighVal = Trim(gReadBuf(4)) 'High
                            
                        '�Ҽ��� ó��, ��� ���� ó��
                        lsEquipRes = strResult
                        strResult = SetResult(strResult, strIntBase)
                        lsResult_Buff = strResult
                        
                        '����
                        strLH = ""
                        If strResult < strLowVal Then
                            strLH = "L"
                        End If
                        If strResult > strHighVal Then
                            strLH = "H"
                        End If
                        
                        '-- Work List
                        SetText vasWorkList, "ó�����", gRow, colState                 '11 �������
                        If strBarNo = "PCTC1026" Then
                            SetText vasWorkList, "QC(Low)", gRow, colPName
                        End If
                        If strBarNo = "PCTC2026" Then
                            SetText vasWorkList, "QC(High)", gRow, colPName
                        End If
                        
                        '-- ������� ��ȸ
                              SQL = " Select PASTREQNO, LABRES "
                        SQL = SQL & "   From PastRes "
                        SQL = SQL & "  Where REQNO = '" & strBarNo & "'"
                        SQL = SQL & "    And ITEMCD = '" & lsExamCode & "'"
                        SQL = SQL & "  Order by PASTREQNO DESC"
                        Res = GetDBSelectColumn(gServer, SQL)
                        
                        If Res > 0 Then
                            strOldRegNo = Trim(gReadBuf(0))     '���� �Ƿڹ�ȣ
                            strOldResult = Trim(gReadBuf(1))    '���� ���
                        End If
                        
                        '-- ��� List
                        SetText vasWorkList, lsResult_Buff, gRow, colrRslt          '���
                        SetText vasWorkList, strDil, gRow, colrDil          'Dilution
                        SetText vasWorkList, strLH, gRow, colrHL               '���� High/Low
                        strLH = ""
                        If IsNumeric(strResult) Then
'                            If Val(strResult) < Val(strLowVal) Then
'                                strLH = "L"
'                            End If
                            If Val(strResult) > Val(strHighVal) Then
                                strLH = "H"
                            End If
                        End If
                        SetText vasWorkList, strOldResult, gRow, colrOldResult               '���� ���
                        SetText vasWorkList, strOldRegNo, gRow, colrOldBar               '���� �Ƿڹ�ȣ
                        SetText vasWorkList, strLowVal, gRow, colrLow               '����ġ ����
                        SetText vasWorkList, strHighVal, gRow, colrHigh               '����ġ ����
                        
                        '-- ���� ����
                        SetLocalDB gRow, gRow, "1", lsEquipRes
                                    
                        lsResult_Buff = ""
                        
                        strState = ""
                            
                        End If
                    End If
                End If
            Case "L"
                '## DB�� �������
                If MnTransAuto.Checked = True And strState = "R" Then
                    
                    Res = SaveTransDataW(gRow)
                    
                    If Res = -1 Then
                        '-- ���� ����
                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                        SetText vasID, "Failed", gRow, colState
                    Else
                        '-- ���� ����
                        SetBackColor vasWorkList, gRow, gRow, 1, colState, 202, 255, 112
                        SetText vasWorkList, "Trans", gRow, colState
                        SetText vasWorkList, "0", gRow, colCheckBox
                        
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
                    strState = ""
                End If
            
                'SetText vasID, "Result", gRow, colState
'                strState = ""
        
        End Select
    Next

End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvDataAU()
    Dim strRcvBuf    As String   '������ Data
    Dim strType      As String   '������ Record Type
    Dim strBarNo     As String   '������ ���ڵ��ȣ
    Dim strSeq       As String   '������ Sequence
    Dim strRackNo    As String   '������ Rack Or Disk No
    Dim strTubePos   As String   '������ Tube Position
    Dim strIntBase   As String   '������ ������ �˻��
    Dim strResult    As String   '������ ���
    Dim strQCResult  As String   '������ ���(QC)
    Dim strFlag      As String   '������ Abnormal Flag
    Dim strComm      As String   '������ Comment
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

                strTmp = Mid$(strRcvBuf, 29)
                                
                Call SetPatInfo(strBarNo)
                
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
                        
                        '-- ���� ���� ���
                        If Res > 0 Then
                            lsExamCode = Trim(gReadBuf(0))
                            lsExamName = Trim(gReadBuf(1))
                            lsSeqNo = Trim(gReadBuf(2))
                            
                            lsResRow = vasRes.DataRowCnt + 1
                            If vasRes.MaxRows < lsResRow Then
                                vasRes.MaxRows = lsResRow
                            End If
                            
                            '�Ҽ��� ó��, ��� ���� ó��
                            lsEquipRes = strResult
                            strResult = SetResult(strResult, strIntBase)
                            lsResult_Buff = strResult
                            
                            '-- Work List
                            SetText vasID, "Result", gRow, colState                 '11 �������
                            

                            '-- ��� List
                            SetText vasRes, strIntBase, lsResRow, colEquipCode      '����ڵ�
                            SetText vasRes, lsExamCode, lsResRow, colExamCode       '�˻��ڵ�
                            SetText vasRes, lsExamName, lsResRow, colExamName       '�˻��
                            SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                            SetText vasRes, strResult, lsResRow, colResult          '���
                            SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                            SetText vasRes, strComm, lsResRow, 7                    'Flag
                            '-- ���� ����
                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
                                        
                            lsResult_Buff = ""
                            
                            strState = "R"
                            
                        '-- ���� ���� ���
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
                                
                                '�Ҽ��� ó��, ��� ���� ó��
                                lsEquipRes = strResult
                                strResult = SetResult(strResult, strIntBase)
                                lsResult_Buff = strResult
                                
                                '-- Work List
                                SetText vasID, "Result", gRow, colState                 '�������
                                
                                '-- ��� List
                                SetText vasRes, strIntBase, lsResRow, colEquipCode      '����ڵ�
                                SetText vasRes, lsExamCode, lsResRow, colExamCode       '�˻��ڵ�
                                SetText vasRes, lsExamName, lsResRow, colExamName       '�˻��
                                SetText vasRes, lsEquipRes, lsResRow, colMachResult     '�����
                                SetText vasRes, strResult, lsResRow, colResult          '���
                                SetText vasRes, lsSeqNo, lsResRow, colSeq               '����
                                SetText vasRes, strComm, lsResRow, colFLAG              'Flag
                                '-- ���� ����
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
                        '-- ���� ����
                        SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                        SetText vasID, "Failed", gRow, colState
                    Else
                        '-- ���� ����
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
' asRow2 = ��� List
Function SetLocalDB(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String, Optional asEquipResult As String = "")
    Dim sCnt As String
    Dim sExamDate As String
    'Dim Rs As ADODB.Recordset
    Dim intSeq As Integer
    
    sExamDate = Format(dtpToday, "yyyymmdd")
    
          SQL = " DELETE FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf
    SQL = SQL & "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & Trim(GetText(vasWorkList, asRow1, colBarcode)) & "' " & vbCrLf
    SQL = SQL & "   AND EXAMCODE = 'X274'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
          SQL = ""
    SQL = SQL & "INSERT INTO PATRESULT("
    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
                "EQUIPRESULT,RESULT,REFFLAG,EXAMNAME,SENDFLAG,EXAMUID, " & vbCrLf
    SQL = SQL & "SAMPLETYPE,WORKNO,WORKKEY,SALETEAM,DEALTEAM,OLDRESULT,OLDBARCODE,REFVALUE,REFLOW,REFHIGH,DELTAVALUE,SENDDATE) " & vbCrLf
    SQL = SQL & "VALUES("
    SQL = SQL & "'" & Trim(Format(dtpToday.Value, "YYYYMMDD")) & "', "
    SQL = SQL & "'" & gEquip & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colBarcode)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colRack)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colPos)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colPID)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colPName)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colSex)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colAge)) & "', "
    SQL = SQL & "'010203', "
    SQL = SQL & "'X274', "
    SQL = SQL & "'', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colrRslt)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colrRslt)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colrDil)) & "', "
    SQL = SQL & "'Procalcintonin', "
    SQL = SQL & "'0', "
    SQL = SQL & "'" & gIFUser & "',"
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colSPC)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colWN)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colWK)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colSale)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colCST)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colrOldResult)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colrOldBar)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colrHL)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colrLow)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colrHigh)) & "', "
    SQL = SQL & "'" & Trim(GetText(vasWorkList, asRow1, colrQC)) & "', "
    If blnSame = True Then
        SQL = SQL & "'" & Format(Now, "yyyymmddhhmmss") & "') "
    Else
        SQL = SQL & "'') "
    End If
    blnSame = False
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If

End Function

' asRow1 = Work List
Function SetLocalDBUP(ByVal asRow1 As Long, ByVal objSpd As Object)
    Dim sCnt As String
    Dim sExamDate As String
    'Dim Rs As ADODB.Recordset
    Dim intSeq As Integer
    
    sExamDate = Format(dtpToday, "yyyymmdd")
    
          SQL = " DELETE FROM PATRESULT " & vbCrLf
    SQL = SQL & " WHERE EXAMDATE = '" & Format(dtpToday, "YYYYMMDD") & "' " & vbCrLf
    SQL = SQL & "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & strPrevRegNo & "' " & vbCrLf
    SQL = SQL & "   AND EXAMCODE = 'X274'"
          
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
          SQL = ""
    SQL = SQL & "INSERT INTO PATRESULT("
    SQL = SQL & "EXAMDATE,EQUIPNO,BARCODE,DISKNO,POSNO," & vbCrLf & _
                "PID,PNAME,PSEX,PAGE,EQUIPCODE,EXAMCODE,SEQNO," & vbCrLf & _
                "EQUIPRESULT,RESULT,REFFLAG,EXAMNAME,SENDFLAG,EXAMUID, " & vbCrLf
    SQL = SQL & "SAMPLETYPE,WORKNO,WORKKEY,SALETEAM,DEALTEAM,OLDRESULT,OLDBARCODE,REFVALUE,REFLOW,REFHIGH,DELTAVALUE) " & vbCrLf
    SQL = SQL & "VALUES("
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colOrdDate)) & "', "
    SQL = SQL & "'" & gEquip & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colBarcode)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colRack)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colPos)) & "', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colPID)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colPName)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colSex)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colAge)) & "', "
    SQL = SQL & "'010203', "
    SQL = SQL & "'X274', "
    SQL = SQL & "'', " & vbCrLf
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colrRslt)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colrRslt)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colrDil)) & "', "
    SQL = SQL & "'Procalcintonin', "
    SQL = SQL & "'0', "
    SQL = SQL & "'" & gIFUser & "',"
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colSPC)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colWN)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colWK)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colSale)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colCST)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colrOldResult)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colrOldBar)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colrHL)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colrLow)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colrHigh)) & "', "
    SQL = SQL & "'" & Trim(GetText(objSpd, asRow1, colrQC)) & "') "
    
    Res = SendQuery(gLocal, SQL)
    
    If Res = -1 Then
        SaveQuery SQL
        Exit Function
    End If

End Function

' asRow1 = Work List
Function SetLocalDBUP_QC(ByVal asRow1 As Long, ByVal objSpd As Object)
    Dim sCnt As String
    Dim sExamDate As String
    'Dim Rs As ADODB.Recordset
    Dim intSeq As Integer
        
          SQL = " UPDATE PATRESULT SET" & vbCrLf
    SQL = SQL & "  DELTAVALUE = '" & Trim(GetText(objSpd, asRow1, colrQC)) & "'"
    SQL = SQL & " WHERE EXAMDATE = '" & Trim(GetText(objSpd, asRow1, colOrdDate)) & "' " & vbCrLf
    SQL = SQL & "   AND EQUIPNO = '" & gEquip & "' " & vbCrLf
    SQL = SQL & "   AND BARCODE = '" & Trim(GetText(objSpd, asRow1, colBarcode)) & "' " & vbCrLf
    SQL = SQL & "   AND EXAMCODE = 'X274'"
          
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
    sMsg = "�˻��ڸ� �Է����ּ���."
    lblUser.Caption = InputBox(sMsg, "�˻��� �Է�")

End Sub





Private Sub txtStartNum_GotFocus()
    
    Call SelectFocus(txtStartNum)
    
End Sub

'Private Sub txtNum_KeyPress(KeyAscii As Integer)
'Dim intRow As Integer
'
'    If KeyAscii = 13 Then
'        With vasWorkList
'            For intRow = .ActiveRow To .DataRowCnt
'                '.Row = intRow
'                '.Col = colCheckBox
'                'If .Value = 1 Then
'                    SetText vasWorkList, txtNum, intRow, colSeqNo
'
'                    txtNum = Val(txtNum) + 1
'
'                'End If
'            Next
'        End With
'    End If
'
'End Sub


Private Sub txtStartNum_KeyPress(KeyAscii As Integer)
Dim intRow As Integer
Dim lngNum As Long

    If KeyAscii = 13 Then
        If IsNumeric(txtStartNum.Text) And IsNumeric(txtStopNum.Text) Then
            With vasWorkList
            
            lngNum = txtStartNum.Text
            
            For intRow = .ActiveRow To .DataRowCnt
                
                Call .SetText(colSeqNo, intRow, lngNum)
                If lngNum = txtStopNum.Text Then
                    Exit For
                End If
                lngNum = lngNum + 1
                
            Next
            
            End With
        Else
            MsgBox "��ȿ�� �Է°��� �ƴմϴ�."
            Exit Sub
        End If
    
'        txtStopNum.SetFocus
    
    End If
    
End Sub

Private Sub txtStopNum_GotFocus()
    
    Call SelectFocus(txtStopNum)
    
End Sub

Private Sub txtStopNum_KeyPress(KeyAscii As Integer)
Dim intRow As Integer
Dim lngNum As Long

    If KeyAscii = 13 Then
        If IsNumeric(txtStartNum.Text) And IsNumeric(txtStopNum.Text) Then
            With vasWorkList
            
            lngNum = txtStartNum.Text
            
            For intRow = .ActiveRow To .DataRowCnt
                
                Call .SetText(colSeqNo, intRow, lngNum)
                If lngNum = txtStopNum.Text Then
                    Exit For
                End If
                lngNum = lngNum + 1
                
            Next
            
            End With
        Else
            MsgBox "��ȿ�� �Է°��� �ƴմϴ�."
        End If
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
    'Local���� �ҷ�����
    ClearSpread vasRes
    
    '����ڵ�, �˻��ڵ�, �˻��, ���, ����
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
'        If MsgBox("�ش� ȯ�ڰ���� �����Ͻðڽ��ϱ�?", vbInformation + vbYesNo, "�˸�") = vbNo Then
'            Exit Sub
'        End If
'
'        SQL = " DELETE FROM PATRESULT " & vbCrLf & _
'              " WHERE EQUIPNO = '" & gEquip & "' " & vbCrLf & _
'              " AND BARCODE = '" & lsID & "' " & vbCrLf & _
'              " AND PID = '" & lsPid & "' " & vbCrLf & _
'              " AND DISKNO = '" & Trim(GetText(vasID, iRow, colCST)) & "' " & vbCrLf & _
'              " AND POSNO = '" & Trim(GetText(vasID, iRow, colSPC)) & "' " & vbCrLf & _
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
'        'Local���� �ҷ�����
'        ClearSpread vasTemp
'
'        '����ڵ�, �˻��ڵ�, �˻��, ���, ����
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
'                  "VALUES('" & gEquip & "', '" & Trim(GetText(vasID, iRow, colBarcode)) & "', '" & Trim(GetText(vasID, iRow, colCST)) & "', " & vbCrLf & _
'                  "'" & Trim(GetText(vasID, iRow, colSPC)) & "', '" & Trim(GetText(vasID, iRow, colPID)) & "', '" & Trim(GetText(vasID, iRow, colPName)) & "', " & vbCrLf & _
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
'                  " AND DISKNO = '" & Trim(GetText(vasID, iRow, colCST)) & "' " & vbCrLf & _
'                  " AND POSNO = '" & Trim(GetText(vasID, iRow, colSPC)) & "' " & vbCrLf & _
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
'                SQL = SQL & vbCrLf & "   AND DISKNO = '" & Trim(GetText(vasID, iRow, colCST)) & "' "
'                SQL = SQL & vbCrLf & "   AND POSNO = '" & Trim(GetText(vasID, iRow, colSPC)) & "' "
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

'Private Sub vasID_KeyPress(KeyAscii As Integer)
'Dim intRow As Integer
'Dim lngNum As Long
'
'    If KeyAscii = 13 Then
'        vasID.Row = vasID.ActiveRow
'        vasID.Col = colSeqNo
'        If Not IsNumeric(vasID.Text) Then
'            Exit Sub
'        End If
'
'        lngNum = vasID.Text
'
'        For intRow = vasID.ActiveRow + 1 To vasID.DataRowCnt
'
'            lngNum = lngNum + 1
'            Call vasID.SetText(colSeqNo, intRow, lngNum)
'
'        Next
'
'        txtNum.Text = lngNum
'
'    End If
'
'End Sub
'
'Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
'    Dim lRow As Long
'
'    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
'        lRow = vasID.ActiveRow
'        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
'
'        vasID_Click colBarcode, lRow
'    End If
'End Sub

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
    'Local���� �ҷ�����
    ClearSpread vasRRes
    
    '����ڵ�, �˻��ڵ�, �˻��, ���, ����
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


Private Sub vasRID_KeyPress(KeyAscii As Integer)

    With vasRID
        If KeyAscii = 13 Then
            .Row = .ActiveRow
            .Col = .ActiveCol
            
            If .ActiveCol = colBarcode Then
                If strPrevRegNo = Trim(.Text) Then
                    Exit Sub
                End If
                If GetSampleInfoS(.ActiveRow) > 0 Then
                    '-- ���� ����
                    Call SetLocalDBUP(.ActiveRow, vasRID)
                End If
            ElseIf .ActiveCol = colrQC Then
                '-- ���� ����
                Call SetLocalDBUP_QC(.ActiveRow, vasRID)
            End If
        End If
    End With

End Sub

Private Sub vasRID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasRID.ActiveRow
        If lRow < 1 Or lRow > vasRID.DataRowCnt Then Exit Sub
            
        vasRID_Click colBarcode, lRow
    End If
End Sub


Private Sub vasRID_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    With vasRID
        If NewCol = colBarcode Then
            .Row = NewRow
            .Col = NewCol
            strPrevRegNo = .Text
        End If
    End With
End Sub

Private Sub vasWorkList_KeyPress(KeyAscii As Integer)

    With vasWorkList
        If KeyAscii = 13 Then
            .Row = .ActiveRow
            .Col = .ActiveCol
            
            If .ActiveCol = colBarcode Then
                If strPrevRegNo = Trim(.Text) Then
                    Exit Sub
                End If
                If GetSampleInfoS(.ActiveRow) > 0 Then
                    '-- ���� ����
                    Call SetLocalDBUP(.ActiveRow, vasRID)
                End If
            ElseIf .ActiveCol = colrQC Then
                '-- ���� ����
                Call SetLocalDBUP_QC(.ActiveRow, vasRID)
            End If
        End If
    End With

End Sub

Private Sub vasWorkList_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    With vasWorkList
        If NewCol = colBarcode Then
            .Row = NewRow
            .Col = NewCol
            strPrevRegNo = .Text
        End If
    End With
End Sub

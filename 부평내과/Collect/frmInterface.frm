VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "BarCode Print"
   ClientHeight    =   11955
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11535
   FillColor       =   &H0000FFFF&
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
   ScaleHeight     =   11955
   ScaleWidth      =   11535
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin VB.PictureBox Picture1 
      Align           =   1  'À§ ¸ÂÃã
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11475
      TabIndex        =   74
      Top             =   0
      Width           =   11535
      Begin MSCommLib.MSComm MSComm1 
         Left            =   10380
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InBufferSize    =   4096
         OutBufferSize   =   1024
         RThreshold      =   1
         RTSEnable       =   -1  'True
         SThreshold      =   1
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   10950
         Top             =   60
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
               Picture         =   "frmInterface.frx":0442
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":09DC
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":0F76
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":1510
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":1DA2
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":1EFC
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":2056
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "BarCode Printer Port"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   180
         Left            =   7830
         TabIndex        =   76
         Top             =   270
         Width           =   2100
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   10080
         Picture         =   "frmInterface.frx":21B0
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Æò¸é
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "Didim BarCode Print Ver 1.0  [ARGOX CP2140]] [ºÎÆò³»°ú]"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   75
         Top             =   180
         Width           =   6255
      End
   End
   Begin TabDlg.SSTab stInterface 
      Height          =   10845
      Left            =   150
      TabIndex        =   26
      Top             =   840
      Width           =   11070
      _ExtentX        =   19526
      _ExtentY        =   19129
      _Version        =   393216
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
      TabCaption(0)   =   "Ãâ·Â"
      TabPicture(0)   =   "frmInterface.frx":273A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPrtSetup"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSetup"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdClear(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdClose(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "ÀçÃâ·Â"
      TabPicture(1)   =   "frmInterface.frx":2756
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdClose(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdClear(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "¿öÅ©¸®½ºÆ®"
      TabPicture(2)   =   "frmInterface.frx":2772
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   10305
         Left            =   -74820
         TabIndex        =   101
         Top             =   360
         Width           =   10725
         Begin VB.CommandButton cmdDelGLU 
            Caption         =   "GLU»èÁ¦"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   9480
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   111
            Top             =   180
            Width           =   1155
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "Ã³¹æ»èÁ¦"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   8250
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   110
            Top             =   180
            Width           =   1155
         End
         Begin VB.CommandButton cmdWorkPrint 
            Caption         =   "Ãâ·Â"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7050
            Style           =   1  '±×·¡ÇÈ
            TabIndex        =   109
            Top             =   180
            Width           =   1155
         End
         Begin VB.CommandButton cmdWorkSearch 
            Caption         =   "Á¶È¸"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5940
            TabIndex        =   104
            Top             =   180
            Width           =   1065
         End
         Begin VB.ComboBox cboPart 
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmInterface.frx":278E
            Left            =   3720
            List            =   "frmInterface.frx":2790
            TabIndex        =   103
            Top             =   210
            Width           =   2025
         End
         Begin VB.CheckBox chkWAll 
            Height          =   315
            Left            =   750
            TabIndex        =   102
            Top             =   750
            Width           =   255
         End
         Begin FPSpread.vaSpread vasWorkPrint 
            Height          =   9465
            Left            =   180
            TabIndex        =   105
            Top             =   660
            Width           =   10365
            _Version        =   393216
            _ExtentX        =   18283
            _ExtentY        =   16695
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
            MaxCols         =   10
            MaxRows         =   20
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmInterface.frx":2792
         End
         Begin MSComCtl2.DTPicker dtpSearch 
            Height          =   375
            Left            =   1020
            TabIndex        =   106
            Top             =   210
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   127401985
            CurrentDate     =   40248
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "°Ë»çÆÄÆ®"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   2850
            TabIndex        =   108
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Á¶È¸ÀÏ"
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   107
            Top             =   300
            Width           =   630
         End
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   -65700
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   100
         Top             =   450
         Width           =   1365
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Á¾·á"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   -65700
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   99
         Top             =   1230
         Width           =   1365
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Á¾·á"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   9330
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   98
         Top             =   2910
         Width           =   1365
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   9330
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   97
         Top             =   2130
         Width           =   1365
      End
      Begin VB.CommandButton cmdSetup 
         Caption         =   "Àåºñ°Ë»çÄÚµå¼³Á¤"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9330
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   96
         Top             =   1380
         Width           =   1365
      End
      Begin VB.CommandButton cmdPrtSetup 
         Caption         =   "¹ÙÄÚµåÇÁ¸°ÅÍ¼³Á¤"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   9330
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   95
         Top             =   600
         Width           =   1365
      End
      Begin VB.Frame Frame6 
         Height          =   10305
         Left            =   -74820
         TabIndex        =   80
         Top             =   360
         Width           =   8835
         Begin VB.CommandButton cmdRePrint 
            Caption         =   "Ãâ·Â"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7350
            TabIndex        =   92
            Top             =   180
            Width           =   1305
         End
         Begin VB.CheckBox chkRePrint 
            Height          =   315
            Left            =   900
            TabIndex        =   82
            Top             =   750
            Width           =   255
         End
         Begin VB.ComboBox cboWorkList 
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmInterface.frx":32CA
            Left            =   5130
            List            =   "frmInterface.frx":32CC
            TabIndex        =   84
            Top             =   210
            Width           =   2025
         End
         Begin VB.CommandButton cmdReSearch 
            Caption         =   "Á¶È¸"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2970
            TabIndex        =   81
            Top             =   210
            Width           =   885
         End
         Begin FPSpread.vaSpread vasRePrint 
            Height          =   9465
            Left            =   180
            TabIndex        =   83
            Top             =   660
            Width           =   8475
            _Version        =   393216
            _ExtentX        =   14949
            _ExtentY        =   16695
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
            SpreadDesigner  =   "frmInterface.frx":32CE
         End
         Begin MSComCtl2.DTPicker dtpBeginDate 
            Height          =   375
            Left            =   1320
            TabIndex        =   87
            Top             =   210
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   127401985
            CurrentDate     =   40248
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "Á¶È¸½ÃÀÛÀÏ"
            Height          =   195
            Index           =   4
            Left            =   210
            TabIndex        =   88
            Top             =   300
            Width           =   1050
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "ÀÇ·ÚÀÏÀÚ"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   4230
            TabIndex        =   85
            Top             =   300
            Width           =   780
         End
      End
      Begin VB.Frame Frame4 
         Height          =   10305
         Left            =   180
         TabIndex        =   53
         Top             =   360
         Width           =   8835
         Begin VB.TextBox txtCnt 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5070
            TabIndex        =   93
            Text            =   "1"
            Top             =   240
            Width           =   435
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Ãâ·Â"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7350
            TabIndex        =   91
            Top             =   180
            Width           =   1305
         End
         Begin VB.Timer Timer1 
            Left            =   5790
            Top             =   210
         End
         Begin VB.CheckBox chkAuto 
            Caption         =   "ÀÚµ¿Ãâ·Â"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   240
            TabIndex        =   86
            Top             =   270
            Value           =   1  'È®ÀÎ
            Width           =   1335
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Á¶È¸"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   56
            Top             =   180
            Width           =   1305
         End
         Begin VB.CheckBox chkPrint 
            Height          =   315
            Left            =   900
            TabIndex        =   54
            Top             =   750
            Width           =   255
         End
         Begin FPSpread.vaSpread vasPrint 
            Height          =   9465
            Left            =   180
            TabIndex        =   55
            Top             =   660
            Width           =   8475
            _Version        =   393216
            _ExtentX        =   14949
            _ExtentY        =   16695
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
            SpreadDesigner  =   "frmInterface.frx":3D6C
         End
         Begin VB.Label Label15 
            Caption         =   "Ãâ·Â¼ö:"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4380
            TabIndex        =   94
            Top             =   300
            Width           =   705
         End
         Begin VB.Label Label14 
            Caption         =   "ÃÊ ÈÄÁ¶È¸"
            Height          =   315
            Left            =   2490
            TabIndex        =   90
            Top             =   300
            Width           =   1245
         End
         Begin VB.Label lblTimer 
            Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  '´ÜÀÏ °íÁ¤
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "±¼¸²Ã¼"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1800
            TabIndex        =   89
            Top             =   210
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         Height          =   9675
         Left            =   -74820
         TabIndex        =   33
         Top             =   360
         Width           =   14625
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
            Left            =   11520
            TabIndex        =   49
            Top             =   240
            Width           =   1395
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
            TabIndex        =   48
            Top             =   240
            Width           =   1395
         End
         Begin VB.CheckBox chkRAll 
            Caption         =   "Check1"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   720
            TabIndex        =   47
            Top             =   780
            Width           =   225
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
            Left            =   3720
            TabIndex        =   45
            Top             =   240
            Width           =   1395
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
            TabIndex        =   44
            Top             =   240
            Width           =   1395
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   7860
            TabIndex        =   38
            Top             =   630
            Width           =   6675
            Begin VB.Label Label11 
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
               TabIndex        =   43
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lblBarcode 
               Caption         =   "1234567890ab"
               Height          =   165
               Index           =   1
               Left            =   1995
               TabIndex        =   42
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label10 
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
               TabIndex        =   41
               Top             =   240
               Width           =   945
            End
            Begin VB.Label lblPname 
               Caption         =   "1234567890ab"
               Height          =   225
               Index           =   1
               Left            =   4590
               TabIndex        =   40
               Top             =   240
               Width           =   1305
            End
            Begin VB.Label lblRrow 
               BackColor       =   &H80000008&
               ForeColor       =   &H8000000E&
               Height          =   315
               Left            =   180
               TabIndex        =   39
               Top             =   720
               Width           =   1155
            End
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
            TabIndex        =   37
            Top             =   210
            Visible         =   0   'False
            Width           =   795
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
            TabIndex        =   36
            Top             =   270
            Value           =   -1  'True
            Width           =   735
         End
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
            TabIndex        =   35
            Top             =   270
            Width           =   735
         End
         Begin FPSpread.vaSpread vasRRes 
            Height          =   8070
            Left            =   7860
            TabIndex        =   34
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
            SpreadDesigner  =   "frmInterface.frx":480A
         End
         Begin MSComCtl2.DTPicker dtpExamDate 
            Height          =   315
            Left            =   1140
            TabIndex        =   46
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
            Format          =   127401984
            CurrentDate     =   40457
         End
         Begin FPSpread.vaSpread vasRID 
            Height          =   8805
            Left            =   165
            TabIndex        =   50
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
            MaxCols         =   12
            MaxRows         =   20
            MoveActiveOnFocus=   0   'False
            RetainSelBlock  =   0   'False
            ScrollBars      =   2
            SpreadDesigner  =   "frmInterface.frx":8573
            UserResize      =   2
         End
         Begin VB.Label Label13 
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
            TabIndex        =   52
            Top             =   390
            Width           =   780
         End
         Begin VB.Label Label12 
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
            TabIndex        =   51
            Top             =   360
            Width           =   780
         End
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   300
         TabIndex        =   27
         Top             =   9120
         Visible         =   0   'False
         Width           =   7515
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmInterface.frx":9093
            Left            =   5160
            List            =   "frmInterface.frx":9095
            TabIndex        =   29
            Top             =   900
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CheckBox Check1 
            Caption         =   "¿À´õÀü¼ÛÈ¯ÀÚ"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6690
            TabIndex        =   28
            Top             =   900
            Visible         =   0   'False
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker dtpToday 
            Height          =   315
            Left            =   1260
            TabIndex        =   30
            Top             =   240
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
            Format          =   127401984
            CurrentDate     =   40457
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Åõ¸í
            Caption         =   "°Ë»ç¼±ÅÃ"
            BeginProperty Font 
               Name            =   "±¼¸²"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   4290
            TabIndex        =   32
            Top             =   990
            Visible         =   0   'False
            Width           =   720
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
            Left            =   330
            TabIndex        =   31
            Top             =   330
            Width           =   780
         End
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         BorderWidth     =   3
         Height          =   615
         Left            =   9330
         Top             =   2130
         Width           =   1365
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0FF&
         BorderWidth     =   3
         Height          =   615
         Left            =   9330
         Top             =   2910
         Width           =   1365
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         Height          =   615
         Left            =   9330
         Top             =   1380
         Width           =   1365
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         Height          =   615
         Left            =   9330
         Top             =   600
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9600
      Left            =   14340
      TabIndex        =   4
      Top             =   1410
      Visible         =   0   'False
      Width           =   18060
      Begin VB.OptionButton optAuto 
         Caption         =   "Auto"
         Height          =   255
         Left            =   2760
         TabIndex        =   73
         Top             =   60
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.OptionButton optManual 
         Caption         =   "Manual"
         Height          =   255
         Left            =   2760
         TabIndex        =   72
         Top             =   390
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmd_Trans 
         Caption         =   "¼±ÅÃÀü¼Û"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4020
         TabIndex        =   71
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Åë½Å¼³Á¤"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5130
         TabIndex        =   70
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtToday 
         Appearance      =   0  'Æò¸é
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1155
         TabIndex        =   68
         Text            =   "2002/02/18"
         Top             =   180
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox ChkAll1 
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   15780
         TabIndex        =   66
         Top             =   510
         Width           =   165
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FF8080&
         Caption         =   "Ãë¼Ò"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   29100
         Style           =   1  '±×·¡ÇÈ
         TabIndex        =   65
         Top             =   1530
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox txtDate 
         Height          =   405
         Left            =   16470
         TabIndex        =   64
         Top             =   930
         Width           =   2325
      End
      Begin VB.CheckBox ChkAll 
         Height          =   315
         Left            =   19110
         TabIndex        =   62
         Top             =   540
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtMsg 
         ForeColor       =   &H000000C0&
         Height          =   825
         Left            =   14190
         MultiLine       =   -1  'True
         ScrollBars      =   2  '¼öÁ÷
         TabIndex        =   59
         Top             =   9510
         Visible         =   0   'False
         Width           =   11745
      End
      Begin VB.TextBox txtAll 
         Height          =   375
         Left            =   15690
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   2580
         Width           =   2055
      End
      Begin VB.TextBox txtTemp 
         Height          =   375
         Left            =   15690
         TabIndex        =   57
         Top             =   1470
         Width           =   2055
      End
      Begin FPSpread.vaSpread vasWork 
         Height          =   6045
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   9195
         _Version        =   393216
         _ExtentX        =   16219
         _ExtentY        =   10663
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   19
         MaxRows         =   100
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15526606
         ShadowDark      =   13815180
         SpreadDesigner  =   "frmInterface.frx":9097
      End
      Begin VB.CommandButton cmdWorkDel 
         Caption         =   "¼±ÅÃ»èÁ¦"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8610
         TabIndex        =   25
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton cmdWorkSave 
         Caption         =   "WORK ÀúÀå"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7170
         TabIndex        =   24
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "¢¹ ºÒ·¯¿À±â"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12510
         TabIndex        =   23
         Top             =   150
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtChartNo 
         Height          =   315
         Left            =   3090
         TabIndex        =   22
         Top             =   5670
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPatNm 
         Height          =   315
         Left            =   990
         TabIndex        =   20
         Top             =   5700
         Visible         =   0   'False
         Width           =   945
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   2415
         Left            =   180
         TabIndex        =   16
         Top             =   7050
         Width           =   9045
         _Version        =   393216
         _ExtentX        =   15954
         _ExtentY        =   4260
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   19
         MaxRows         =   100
         Protect         =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   14680063
         ShadowDark      =   13815180
         SpreadDesigner  =   "frmInterface.frx":A839
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   8775
         Left            =   7200
         TabIndex        =   18
         Top             =   690
         Visible         =   0   'False
         Width           =   7395
         _Version        =   393216
         _ExtentX        =   13044
         _ExtentY        =   15478
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   19
         MaxRows         =   100
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   15323903
         ShadowDark      =   13815180
         SpreadDesigner  =   "frmInterface.frx":C01A
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "¢¹¢¹µî·Ï"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5940
         TabIndex        =   17
         Top             =   210
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Height          =   585
         Left            =   90
         TabIndex        =   8
         Top             =   8940
         Visible         =   0   'False
         Width           =   3555
         Begin VB.TextBox txtRack 
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   1050
            TabIndex        =   10
            Text            =   "0"
            Top             =   195
            Width           =   675
         End
         Begin VB.TextBox txtPos 
            Appearance      =   0  'Æò¸é
            BackColor       =   &H00FFFFFF&
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
            Height          =   285
            Left            =   2790
            TabIndex        =   9
            Text            =   "1"
            Top             =   195
            Width           =   675
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   315
            Left            =   90
            TabIndex        =   11
            Top             =   180
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Disk"
            BackColor       =   15526606
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSPanel SSPanel5 
            Height          =   315
            Left            =   1830
            TabIndex        =   12
            Top             =   180
            Width           =   915
            _Version        =   65536
            _ExtentX        =   1614
            _ExtentY        =   556
            _StockProps     =   15
            Caption         =   "Pos"
            BackColor       =   15526606
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±¼¸²Ã¼"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.CommandButton cmdCall 
         Caption         =   "ÀúÀåµ¥ÀÌÅÍ ºÒ·¯¿À±â"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4650
         TabIndex        =   5
         Top             =   5640
         Visible         =   0   'False
         Width           =   2385
      End
      Begin MSComCtl2.DTPicker dtpStopDt 
         Height          =   345
         Left            =   3060
         TabIndex        =   13
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   127401985
         CurrentDate     =   40248
      End
      Begin MSComCtl2.DTPicker dtpStartDt 
         Height          =   375
         Left            =   1140
         TabIndex        =   14
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   127401985
         CurrentDate     =   40248
      End
      Begin Threed.SSPanel sspMode 
         Height          =   675
         Left            =   16410
         TabIndex        =   61
         Top             =   180
         Visible         =   0   'False
         Width           =   825
         _Version        =   65536
         _ExtentX        =   1455
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "Àü¼Û¸ðµå"
         ForeColor       =   16777215
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   3
         BorderWidth     =   5
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   4455
         Left            =   15570
         TabIndex        =   63
         Top             =   1710
         Width           =   3555
         _Version        =   393216
         _ExtentX        =   6271
         _ExtentY        =   7858
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpreadDesigner  =   "frmInterface.frx":D7BC
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   3885
         Left            =   15450
         TabIndex        =   67
         Top             =   5220
         Width           =   5325
         _Version        =   393216
         _ExtentX        =   9393
         _ExtentY        =   6853
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
         SpreadDesigner  =   "frmInterface.frx":11D15
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   675
         Left            =   60
         TabIndex        =   77
         Top             =   720
         Width           =   4170
         _Version        =   65536
         _ExtentX        =   7355
         _ExtentY        =   1191
         _StockProps     =   15
         Caption         =   "     BarCode Print"
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²Ã¼"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         FloodColor      =   14737632
         Alignment       =   1
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   465
            Left            =   6540
            TabIndex        =   79
            Top             =   60
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtBuffer 
            Height          =   285
            Left            =   5430
            TabIndex        =   78
            Top             =   150
            Visible         =   0   'False
            Width           =   1125
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "°Ë»çÀÏÀÚ"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   69
         Top             =   240
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "~"
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
         Left            =   15900
         TabIndex        =   60
         Top             =   1260
         Width           =   120
      End
      Begin VB.Label Label8 
         Caption         =   "Ã­Æ®¹øÈ£"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2130
         TabIndex        =   21
         Top             =   5730
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "È¯ÀÚ¸í"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   5760
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "ÀÇ·ÚÀÏÀÚ"
         BeginProperty Font 
            Name            =   "±¼¸²Ã¼"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Åõ¸í
         Caption         =   "~"
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
         Left            =   2850
         TabIndex        =   7
         Top             =   330
         Width           =   120
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¿À´Ã °Ë»ç¼ö"
      Height          =   195
      Left            =   16155
      TabIndex        =   3
      Top             =   1095
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblToday 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "100,000"
      Height          =   195
      Left            =   17415
      TabIndex        =   2
      Top             =   1095
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÇöÀç °Ë»ç¼ö"
      Height          =   195
      Left            =   16590
      TabIndex        =   1
      Top             =   1470
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblCurrent 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "0"
      Height          =   195
      Left            =   18045
      TabIndex        =   0
      Top             =   1425
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Menu mnuPop 
      Caption         =   "pp"
      Visible         =   0   'False
      Begin VB.Menu subUp 
         Caption         =   "°ËÃ¼¹øÈ£ ¼öÁ¤"
      End
      Begin VB.Menu subDel 
         Caption         =   "»èÁ¦"
      End
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const colCheckBox = 1
Const colRack = 2
Const ColPos = 3
Const colSampleNo = 4
Const colReceNo = 5    '°ËÃ¼¹øÈ£
Const colPID = 6
Const colPName = 7
'Const colJumin = 8
Const colPSex = 8
Const colPAge = 9
Const colOCnt = 10
'Const colRCnt = 11
Const colState = 11
Const colJumin = 12
Const colReqDate = 13   'Á¢¼öÀÏÀÚ

Const colEquipExam = 3
Const colExamCode = 4
Const colExamName = 5
Const colResult = 6
Const colRCheck = 7
Const colPCheck = 8
Const colDCheck = 9
Const colUnit = 10
Const colRef = 11
Const colPanic = 12
Const colResult1 = 13
Const colResState = 14

Dim ConfirmData As String
Dim aCount

Public gPID As String
Public gTestID As String
Public gSpecID As String
Public glRow As Long
Public gCount As String
Public gOCnt As Integer
Public gOCnt_1 As Integer
Public gRCnt As Integer
Public gCheck As String

Dim iIndex As Integer

Dim MyBuff As String

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
Dim strBufferData As String

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

Dim blnSameRecord As Boolean
'Dim intSelRow As Integer
Dim OrderSort_Flag As Integer
Dim lngRefresh As Long


Private Sub cboWorkList_Click()
    Call cmdLoad_Click
End Sub

Private Sub Check2_Click()

End Sub

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If ChkAll.Value = 1 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 1
        Next iRow
    ElseIf ChkAll.Value = 0 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 0
        Next iRow
    End If
End Sub

Private Sub ChkAll1_Click()
    Dim iRow As Integer
    
    If ChkAll1.Value = 1 Then
        For iRow = 1 To vasWork.DataRowCnt
            vasWork.Row = iRow
            vasWork.Col = 1
            
            vasWork.Value = 1
        Next iRow
    ElseIf ChkAll1.Value = 0 Then
        For iRow = 1 To vasWork.DataRowCnt
            vasWork.Row = iRow
            vasWork.Col = 1
            
            vasWork.Value = 0
        Next iRow
    End If
End Sub

Private Sub chkPrint_Click()
    Dim iRow As Integer
    
    If chkPrint.Value = 1 Then
        For iRow = 1 To vasPrint.DataRowCnt
            vasPrint.Row = iRow
            vasPrint.Col = 1
            
            vasPrint.Value = 1
        Next iRow
    ElseIf chkPrint.Value = 0 Then
        For iRow = 1 To vasPrint.DataRowCnt
            vasPrint.Row = iRow
            vasPrint.Col = 1
            
            vasPrint.Value = 0
        Next iRow
    End If
End Sub

Private Sub chkRePrint_Click()
    Dim iRow As Integer
    
    If chkRePrint.Value = 1 Then
        For iRow = 1 To vasRePrint.DataRowCnt
            vasRePrint.Row = iRow
            vasRePrint.Col = 1
            
            vasRePrint.Value = 1
        Next iRow
    ElseIf chkRePrint.Value = 0 Then
        For iRow = 1 To vasRePrint.DataRowCnt
            vasRePrint.Row = iRow
            vasRePrint.Col = 1
            
            vasRePrint.Value = 0
        Next iRow
    End If

End Sub

Private Sub chkWAll_Click()
    Dim iRow As Integer
    
    If chkWAll.Value = 1 Then
        For iRow = 1 To vasWorkPrint.DataRowCnt
            vasWorkPrint.Row = iRow
            vasWorkPrint.Col = 1
            
            vasWorkPrint.Value = 1
        Next iRow
    ElseIf chkWAll.Value = 0 Then
        For iRow = 1 To vasWorkPrint.DataRowCnt
            vasWorkPrint.Row = iRow
            vasWorkPrint.Col = 1
            
            vasWorkPrint.Value = 0
        Next iRow
    End If

End Sub

Private Sub cmd_Trans_Click()
'¼±ÅÃÀü¼Û
    Dim vasIDRow As Integer
    Dim vasResRow As Integer
    Dim iRow As Integer
    Dim liRet As Integer
    Dim FindFile As String

    If MsgBox(" " & vbCrLf & "¼±ÅÃÀü¼ÛÀ» ÇÏ½Ã°Ú½À´Ï±î?" & vbCrLf & " ", vbInformation + vbOKCancel, "¾Ë¸²:¼±ÅÃÀü¼Û") = vbCancel Then
        Exit Sub
    End If
    
    For vasIDRow = 1 To vasList.DataRowCnt
        vasList.Col = 1
        vasList.Row = vasIDRow
        If vasList.Value = 1 Then
            liRet = -1
            If Trim(GetText(vasList, vasIDRow, 4)) <> "" Then
'                liRet = Make_XML_File(vasIDRow)
                liRet = Make_XML(vasIDRow)
            End If
            
            If liRet = 1 Then
                SetBackColor vasList, vasIDRow, vasIDRow, colCheckBox, 12, 202, 255, 112
                'SetText vasList, "Àü¼Û¿Ï·á", vasIDRow, colState
                
                FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_In.xml")
                If FindFile <> "" Then
                    Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"     'Àü¼Û¿Ï·á°¡ µÆÀ»¶§ ÆÄÀÏÁö¿ì±â
                End If
                      
                      SQL = " Update pat_res Set "
                SQL = SQL & " TransYN = 'Y', "
                SQL = SQL & " TransDt = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' "
                vasList.Row = vasIDRow: vasList.Col = 4
                SQL = SQL & " Where ChartNo  = '" & Trim(vasList.Text) & "' "
                vasList.Row = vasIDRow: vasList.Col = 12
                SQL = SQL & "   and ExamID   = '" & Trim(vasList.Text) & "' "
                vasList.Row = vasIDRow: vasList.Col = 10
                SQL = SQL & "   and CommDate = '" & Trim(vasList.Text) & "'"
                Res = SendQuery(gLocal, SQL)
                
            Else
                SetBackColor vasID, vasIDRow, vasIDRow, colCheckBox, 12, 255, 0, 0
                'SetText vasID, "½ÇÆÐ", vasIDRow, colState
            End If
            vasList.Col = 1
            vasList.Value = "0"
        Else
        
        End If
    Next vasIDRow
    
    If XmlTxtHead = "" Then
        XmlTxtHead = "<?xml version=""1.0"" encoding=""euc-kr""?>" & vbCrLf & _
                     "<?xml-stylesheet type=""text/xsl"" href=C:\UBCare\SINAI\IF\Form\ExamIF_Form_05.xsl""?>" & vbCrLf & "<UBCare°Ë»çÁ¤º¸>"
    End If
    
    If XmlTxtTail = "" Then
        XmlTxtTail = "</UBCare°Ë»çÁ¤º¸>"
    End If
    
'    XMLAllTxt = XmlTxtHead & XMLAllTxt & XmlTxtTail
    SaveXMLFile XMLAllTxt
    
End Sub

Function Make_XML_File(asRow) As Integer
    Dim FilNum
    Dim FilNum2
    Dim TxtString As String
    Dim ResultString As String
    Dim TxtRece As String
    Dim i As Long
    Dim PChartNum As String
    Dim PName As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAge As String
    Dim pSex As String
    Dim STxt, NumTxt As Long
    Dim SQL As String
    Dim PEquipno As String
    
    Dim PExamname As String
    Dim PEquipCode As String
    Dim j As Long
    Dim BarFlag As Integer
    Dim pResult As String
    Dim pExamdate As String
    Dim pOpinion As String
    Dim TxtPat As String
    Dim IOGubun As String
    Dim TestNum As String
    
    Make_XML_File = -1

    ClearSpread vasResTemp
    
    SQL = "select  pid, examcode, recedate, barcode,pname, pjumin, examdate, gubun, subcode,result " & vbCrLf & _
          "from pat_res where examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & CR & _
          " and result <> '' " & CR & _
          " And equipno = '" & gEquip & "' and barcode = '" & Trim(GetText(vasID, asRow, colReceNo)) & "'"
    Res = db_select_Vas(gLocal, SQL, vasResTemp)

    
    For i = 1 To vasResTemp.DataRowCnt
'    XMLAllTxt = ""
        PID = Trim(GetText(vasResTemp, i, 1))
        PExamCode = Trim(GetText(vasResTemp, i, 2))
        PReceDate = Trim(GetText(vasResTemp, i, 3))
        PChartNum = Trim(GetText(vasResTemp, i, 4))
        PName = Trim(GetText(vasResTemp, i, 5))
        PJumin = Mid(Trim(GetText(vasResTemp, i, 6)), 1, 6) & "-" & Mid(Trim(GetText(vasResTemp, i, 6)), 7)
        pExamdate = Trim(GetText(vasResTemp, i, 7))
        IOGubun = Trim(GetText(vasResTemp, i, 8))
        TestNum = Trim(GetText(vasResTemp, i, 9))
        pResult = Trim(GetText(vasResTemp, i, 10))
        XMLAllTxt = XMLAllTxt & "<°Ë»ç><¾÷Ã¼>ACK</¾÷Ã¼><¿ä¾ç±â°ü¹øÈ£>11397730</¿ä¾ç±â°ü¹øÈ£><Â÷Æ®¹øÈ£>" & PChartNum & "</Â÷Æ®¹øÈ£><¼öÁøÀÚ¸í>" & PName & "</¼öÁøÀÚ¸í><ÁÖ¹Îµî·Ï¹øÈ£>" & PJumin & "</ÁÖ¹Îµî·Ï¹øÈ£><³»¿ø¹øÈ£>" & PID & "</³»¿ø¹øÈ£><ÀÇ·ÚÀÏ>" & PReceDate & "</ÀÇ·ÚÀÏ><°Ë»ç¹øÈ£>" & TestNum & "</°Ë»ç¹øÈ£><°Ë»çID>" & PExamCode & "</°Ë»çID><¾÷Ã¼°Ë»çID></¾÷Ã¼°Ë»çID><°ËÃ¼></°ËÃ¼><°á°úÄ¡>" & pResult & "</°á°úÄ¡><ÂüÁ¶Ä¡></ÂüÁ¶Ä¡><¼Ò°ß></¼Ò°ß><°á°úÀÏ>" & pExamdate & "</°á°úÀÏ><ÀÔ¿ø¿Ü·¡±¸ºÐ>" & IOGubun & "</ÀÔ¿ø¿Ü·¡±¸ºÐ></°Ë»ç>"
    Next
    
    Make_XML_File = 1
    
End Function


Function Make_XML(asRow) As Integer
Dim varTmp As Variant
Dim strTmp As String
Dim strRslt As String

    With vasList
        .Row = asRow
                    XMLAllTxt = XMLAllTxt & "<°Ë»ç>"
        .Col = 2:   XMLAllTxt = XMLAllTxt & "<¾÷Ã¼>" & Trim(.Text) & "</¾÷Ã¼>"
        .Col = 3:   XMLAllTxt = XMLAllTxt & "<¿ä¾ç±â°ü¹øÈ£>" & Trim(.Text) & "</¿ä¾ç±â°ü¹øÈ£>"
'        .Col = 3:   XMLAllTxt = XMLAllTxt & "<¿ä¾ç±â°ü¹øÈ£>32316577</¿ä¾ç±â°ü¹øÈ£>"
        .Col = 4:   XMLAllTxt = XMLAllTxt & "<Â÷Æ®¹øÈ£>" & Trim(.Text) & "</Â÷Æ®¹øÈ£>"
        .Col = 5:   XMLAllTxt = XMLAllTxt & "<¼öÁøÀÚ¸í>" & Trim(.Text) & "</¼öÁøÀÚ¸í>"
        .Col = 8:   XMLAllTxt = XMLAllTxt & "<ÁÖ¹Îµî·Ï¹øÈ£>" & Trim(.Text) & "</ÁÖ¹Îµî·Ï¹øÈ£>"
        .Col = 9:   XMLAllTxt = XMLAllTxt & "<³»¿ø¹øÈ£>" & Trim(.Text) & "</³»¿ø¹øÈ£>"
        .Col = 10:  XMLAllTxt = XMLAllTxt & "<ÀÇ·ÚÀÏ>" & Trim(.Text) & "</ÀÇ·ÚÀÏ>"
        .Col = 11:  XMLAllTxt = XMLAllTxt & "<°Ë»ç¹øÈ£>" & Trim(.Text) & "</°Ë»ç¹øÈ£>"
        .Col = 12:  XMLAllTxt = XMLAllTxt & "<°Ë»çID>" & Trim(.Text) & "</°Ë»çID>"
        .Col = 13:  XMLAllTxt = XMLAllTxt & "<¾÷Ã¼°Ë»çID>" & Trim(.Text) & "</¾÷Ã¼°Ë»çID>"
        .Col = 14:  XMLAllTxt = XMLAllTxt & "<°ËÃ¼>" & Trim(.Text) & "</°ËÃ¼>"
        .Col = 15:  strRslt = Trim(.Text)
                    XMLAllTxt = XMLAllTxt & "<°á°úÄ¡>" & strRslt & "</°á°úÄ¡>"
        .Col = 16:  XMLAllTxt = XMLAllTxt & "<ÂüÁ¶Ä¡>" & Trim(.Text) & "</ÂüÁ¶Ä¡>"
'        .Col = 17:  XMLAllTxt = XMLAllTxt & "<¼Ò°ß>" & "CCP: " & strRslt & "</¼Ò°ß>"
        .Col = 17:  XMLAllTxt = XMLAllTxt & "<¼Ò°ß>" & Trim(.Text) & "</¼Ò°ß>"
        .Col = 18:  XMLAllTxt = XMLAllTxt & "<°á°úÀÏ>" & Trim(.Text) & "</°á°úÀÏ>"
        .Col = 19:  XMLAllTxt = XMLAllTxt & "<ÀÔ¿ø¿Ü·¡±¸ºÐ>" & Trim(.Text) & "</ÀÔ¿ø¿Ü·¡±¸ºÐ>"
                    XMLAllTxt = XMLAllTxt & "</°Ë»ç>"

    End With
    
    Make_XML = 1
    
End Function

Function Insert_Data(ByVal argSpcRow As Integer) As Integer
'¼­¹öÀÇ µ¥ÀÌÅ¸ º£ÀÌ½º¿¡ ÀúÀå
    Dim iRow As Integer
    Dim i As Integer
    Dim sGubun  As String
    Dim sPID As String
    Dim sDate As String
    
    Dim sNumber As String
    Dim sPan As String
    Dim sLow As String  'ÃÖÀú°ª
    Dim sHigh As String  'ÃÖ´ë°ª
    Dim sResult As String  '°á°úÄ¡
    
    Dim sState As String
    
    Dim iCnt As Integer
    Dim sResDate As String
    Dim sResTime As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PChartNum As String
    Dim PName As String
    Dim PJumin As String
    Dim pExamdate As String
    Dim IOGubun As String
    Dim TestNum As String
    Dim pResult As String
    Dim lsState, lsOrdCnt, lsUID As String
    
    sResDate = Format(Date, "yyyymmdd")
    sResTime = Format(Time, "hhnnss")
    iCnt = 0
    
    sGubun = Mid(Trim(GetText(vasID, argSpcRow, colReceNo)), 1, 1)
    sNumber = Mid(Trim(GetText(vasID, argSpcRow, colReceNo)), 2)
    
    Insert_Data = -1
    
    'Local¿¡¼­ È¯ÀÚº°·Î °á°ú°ª °¡Á®¿À±â
    ClearSpread vasResTemp

    SQL = " Select barcode, equipcode, examcode, result, refflag, panicflag, deltaflag, recedate " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where examdate = '" & Format(Trim(txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And diskno = '" & Trim(GetText(vasID, argSpcRow, colRack)) & "' " & vbCrLf & _
          " And posno = '" & Trim(GetText(vasID, argSpcRow, ColPos)) & "' " & vbCrLf & _
          " And receno = '" & Trim(GetText(vasID, argSpcRow, colSampleNo)) & "' " & vbCrLf & _
          " And barcode = '" & Trim(GetText(vasID, argSpcRow, colReceNo)) & "' " & vbCrLf & _
          " And sendflag = 'A' "
    Res = db_select_Vas(gLocal, SQL, vasResTemp)
    Save_Raw_Data "(" & vasResTemp.DataRowCnt & ")" & SQL
  
    If sGubun = "A" Then
        sDate = Trim(GetText(vasID, argSpcRow, colReqDate))
        If sDate = "" Then
            sDate = Format(txtToday.Text, "yyyymmdd")
        End If
        '¼­¹ö·Î °á°ú°ª ÀúÀåÇÏ±â

        ClearSpread vasResTemp
        SQL = "select  barcode, examcode, recedate, pid,pname, pjumin, examdate, gubun, subcode,result " & vbCrLf & _
              "from pat_res where examdate = '" & Format(CDate(txtToday.Text), "yyyymmdd") & "' " & CR & _
              " And equipno = '" & gEquip & "' and barcode = '" & Trim(GetText(vasID, argSpcRow, colPID)) & "' and pid = '" & Trim(GetText(vasID, argSpcRow, colReceNo)) & "'"
        Res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    For i = 1 To vasResTemp.DataRowCnt
'    XMLAllTxt = ""
        PID = Trim(GetText(vasResTemp, i, 1))
        PExamCode = Trim(GetText(vasResTemp, i, 2))
        PReceDate = Trim(GetText(vasResTemp, i, 3))
        PChartNum = Trim(GetText(vasResTemp, i, 4))
        PName = Trim(GetText(vasResTemp, i, 5))
        PJumin = Mid(Trim(GetText(vasResTemp, i, 6)), 1, 6) & "-" & Mid(Trim(GetText(vasResTemp, i, 6)), 7)
        pExamdate = Trim(GetText(vasResTemp, i, 7))
        IOGubun = Trim(GetText(vasResTemp, i, 8))
        TestNum = Trim(GetText(vasResTemp, i, 9))
        pResult = Trim(GetText(vasResTemp, i, 10))
        XMLAllTxt = XMLAllTxt & "<°Ë»ç><¾÷Ã¼>ACK</¾÷Ã¼><¿ä¾ç±â°ü¹øÈ£>21342784</¿ä¾ç±â°ü¹øÈ£><Â÷Æ®¹øÈ£>" & PID & "</Â÷Æ®¹øÈ£><¼öÁøÀÚ¸í>" & PName & "</¼öÁøÀÚ¸í><ÁÖ¹Îµî·Ï¹øÈ£>" & PJumin & "</ÁÖ¹Îµî·Ï¹øÈ£><³»¿ø¹øÈ£>" & PChartNum & "</³»¿ø¹øÈ£><ÀÇ·ÚÀÏ>" & PReceDate & "</ÀÇ·ÚÀÏ><°Ë»ç¹øÈ£>" & TestNum & "</°Ë»ç¹øÈ£><°Ë»çID>" & PExamCode & "</°Ë»çID><¾÷Ã¼°Ë»çID></¾÷Ã¼°Ë»çID><°ËÃ¼></°ËÃ¼><°á°úÄ¡>" & pResult & "</°á°úÄ¡><ÂüÁ¶Ä¡></ÂüÁ¶Ä¡><¼Ò°ß></¼Ò°ß><°á°úÀÏ>" & pExamdate & "</°á°úÀÏ><ÀÔ¿ø¿Ü·¡±¸ºÐ>" & IOGubun & "</ÀÔ¿ø¿Ü·¡±¸ºÐ></°Ë»ç>"
    Next

    ElseIf sGubun = "B" Then
        sDate = Mid(sResDate, 2, 8)
        sDate = Left(sResDate, 4) & "-" & Mid(sResDate, 5, 2) & "-" & Mid(sResDate, 7, 2)
        'sNumber = Mid(sNumber, 11)

'        sDate = Trim(GetText(vasID, argSpcRow, colReqDate))
'        sNumber = Trim(GetText(vasID, argSpcRow, colReceNo))
        
       'Çö³»°ú °ËÁøÃß°¡
        For i = 1 To vasResTemp.DataRowCnt
            sPan = ""
            
            sResult = Trim(GetText(vasResTemp, i, 4))
            
            'If IsNumeric(sResult) Then
            SQL = "Select res1_gum_code, res1_low, res1_high " _
                & "From mresult001 " _
                    & "Where res1_date = '" & Trim(GetText(vasResTemp, i, 8)) & "' " _
                    & "  And res1_gumno = '" & sNumber & "' " & _
                      "  And res1_gum_code = '" & Trim(GetText(vasResTemp, i, 3)) & "' "
            Res = db_select_Col(gServer, SQL)
            Save_Raw_Data Res & " : " & SQL
            If Trim(gReadBuf(0)) = Trim(GetText(vasResTemp, i, 3)) Then
                sLow = Trim(gReadBuf(1))
                sHigh = Trim(gReadBuf(2))
                
                If IsNumeric(Trim(GetText(vasResTemp, i, 4))) Then
                    If IsNumeric(sLow) Then
                        If CCur(sLow) > CCur(Trim(GetText(vasResTemp, i, 4))) Then
                            sPan = "L"
                        End If
                    End If
                    If IsNumeric(sHigh) Then
                        If CCur(sHigh) < CCur(Trim(GetText(vasResTemp, i, 4))) Then
                            sPan = "H"
                        End If
                    End If
                End If
                Save_Raw_Data Trim(GetText(vasResTemp, i, 3)) & " : " & sLow & " ~ " & sHigh & " => " & sPan
            End If
            'End If
                        
'            If gReadBuf(0) <> "" Then
'                sTmpPan = Trim(gReadBuf(0))
                
                SQL = "Update mresult001 " _
                    & "Set res1_result = '" & Trim(GetText(vasResTemp, i, 4)) & "', " _
                    & "res1_rfm_name = '" & Trim(GetText(vasResTemp, i, 4)) & "', " _
                    & "res1_pan = '" & sPan & "' " _
                    & "Where res1_date = '" & Trim(GetText(vasResTemp, i, 8)) & "' " _
                    & "  And res1_gumno = " & sNumber & _
                      "  And res1_gum_code = '" & Trim(GetText(vasResTemp, i, 3)) & "' "
                
                Res = SendQuery(gServer, SQL)
                Save_Raw_Data Res & " : " & SQL
                If Res = -1 Then
                    Exit Function
                End If
'            End If
        Next i
    End If
    
    Insert_Data = 1
    
End Function

'Private Sub cmdCalendar_Click(Index As Integer)
'    iIndex = Index
'    If Index = 0 Then
'        monvCal.Left = dtpSDate.Left
'        monvCal.Top = 1500
'        monvCal.Visible = True
'
'        monvCal.Value = dtpSDate.Text
'    ElseIf Index = 1 Then
'        monvCal.Left = dtpEDate.Left
'        monvCal.Top = 1500
'        monvCal.Visible = True
'
'        monvCal.Value = dtpEDate.Text
'    End If
'    'monvCal.Visible = True
'End Sub

Private Sub cmdCall_Click()
'    Dim lRow As Long
'
'    ClearSpread vasID
'
'    SQL = " Select diskno, posno, receno, barcode, pid, pname, psex, page, count(*), '', '' " & vbCrLf & _
'          " From pat_res " & vbCrLf & _
'          " Where examdate = '" & Format(txtToday.Text, "yyyymmdd") & "' " & vbCrLf & _
'          " and receno <> '' " & vbCrLf & _
'          " Group By diskno, posno, barcode, pid, pname, '', psex, page,  receno, recedate "
'    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
'    vasSort vasID, 4
'    For lRow = 1 To vasID.DataRowCnt
'        SQL = "select state from Worklist " & vbCrLf & _
'              "WHERE examdate = '" & Format(CDate(frmInterface.txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
'              "  AND SampleID = '" & Trim(GetText(vasID, lRow, colReceNo)) & "' "
'        res = db_select_Col(gLocal, SQL)
'        'Debug.Print Trim(GetText(vasID, lRow, 4))
'        Select Case Trim(gReadBuf(0))
'        Case "A"
'            SetBackColor vasID, lRow, lRow, 1, 1, 255, 255, 112
'        Case "B", "C"
'            SetBackColor vasID, lRow, lRow, 1, 1, 202, 255, 112
'        Case Else
'            SetBackColor vasID, lRow, lRow, 1, 1, 255, 255, 255
'        End Select
'    Next lRow
    Dim PJumin As String
    Dim pGrid_Point As Integer
    Dim adoRS   As New ADODB.Recordset
    
    
          SQL = " Select Company,HospCode,ChartNo,PatName,PatSex,PatAge,PatJumin,PatNo,CommDate,ExamNo,ExamID,ComExamID, "
    SQL = SQL & "        Specimen,Result,Reference,Remark,RsltDate,IOFlag "
    SQL = SQL & "   from pat_res "
'    SQL = SQL & "  where TransDT = '" & Format(txtToday.Text, "yyyymmdd") & "' "
'    SQL = SQL & "  where TransDT Between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "' "
    SQL = SQL & "  where Commdate Between '" & Format(dtpStartDt.Value, "yyyymmdd") & "' and '" & Format(dtpStopDt.Value, "yyyymmdd") & "' "
    
    If Trim(txtChartNo.Text) <> "" Then
        SQL = SQL & "    and ChartNo like '%" & Trim(txtChartNo.Text) & "%' "
    End If
    
    If Trim(txtPatNm.Text) <> "" Then
        SQL = SQL & "    and PatName like '%" & Trim(txtPatNm.Text) & "%' "
    End If
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open SQL, cn
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    With vasID
        .MaxRows = adoRS.RecordCount
        Do While Not adoRS.EOF
            pGrid_Point = pGrid_Point + 1
            .SetText 2, pGrid_Point, adoRS.Fields(0).Value & ""
            .SetText 3, pGrid_Point, adoRS.Fields(1).Value & ""
            .SetText 4, pGrid_Point, adoRS.Fields(2).Value & ""
            .SetText 5, pGrid_Point, adoRS.Fields(3).Value & ""
            .SetText 6, pGrid_Point, adoRS.Fields(4).Value & ""
            .SetText 7, pGrid_Point, adoRS.Fields(5).Value & ""
            .SetText 8, pGrid_Point, adoRS.Fields(6).Value & ""
            .SetText 9, pGrid_Point, adoRS.Fields(7).Value & ""
            .SetText 10, pGrid_Point, adoRS.Fields(8).Value & ""
            .SetText 11, pGrid_Point, adoRS.Fields(9).Value & ""
            .SetText 12, pGrid_Point, adoRS.Fields(10).Value & ""
            .SetText 13, pGrid_Point, adoRS.Fields(11).Value & ""
            .SetText 14, pGrid_Point, adoRS.Fields(12).Value & ""
            .SetText 15, pGrid_Point, adoRS.Fields(13).Value & ""
            .SetText 16, pGrid_Point, adoRS.Fields(14).Value & ""
            .SetText 17, pGrid_Point, adoRS.Fields(15).Value & ""
            .SetText 18, pGrid_Point, adoRS.Fields(16).Value & ""
            .SetText 19, pGrid_Point, adoRS.Fields(17).Value & ""
            adoRS.MoveNext
        Loop
End With
        
End Sub

Private Sub cmdClear_Click(Index As Integer)
Dim iRow As Integer

    If Index = 0 Then
        txtMsg.Text = ""
        txtBuffer = ""
        txtChartNo.Text = ""
        txtPatNm.Text = ""
        
        ClearSpread vasID, 1, 1
        vasID.MaxRows = 1
    
        For iRow = 1 To vasID.DataRowCnt
            vasID.Row = iRow
            vasID.Col = 1
            
            If vasID.Value = 1 Then
                vasDeleteRow vasID, iRow
                
                iRow = iRow - 1
            End If
        Next iRow
        
        ClearSpread vasPrint
        
        ClearSpread vasWork
    '    vasWork.MaxRows = 1
    
        ClearSpread vasList
    '    vasID.MaxRows = 1
    Else
        
        ClearSpread vasRePrint
    
    End If

End Sub

Private Sub cmdClose_Click(Index As Integer)
    
    Unload Me
    
End Sub

Private Sub cmdConfig_Click()
'    frmConfig.SSPanel_machine.Caption = "PhD"
    frmConfig.Show 1
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
    Screen.MousePointer = 0

    Exit Function
        
ErrorTrap:
    
    blnSameRecord = False
    Screen.MousePointer = 0
    
    
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
                SeqSearch = sCnt
                .Action = ActionActiveCell
                .Refresh
                Exit For
            End If
        Next sCnt
    End With

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

Private Sub cmdDel_Click()
    Dim intRow As Integer
    Dim strCommDate As String
    Dim strExamtype As String
    Dim strBarCode  As String
    
    If MsgBox("°Ë»ç°á°ú±îÁö »èÁ¦µË´Ï´Ù. »èÁ¦ÇÏ½Ã°Ú½À´Ï±î?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
        With vasWorkPrint
            For intRow = 1 To .DataRowCnt
                .Row = intRow
                .Col = 1
                If .Value = "1" Then
                    .Col = 2: strCommDate = Format(Trim(.Text), "yyyymmdd") '-- ÀÇ·ÚÀÏÀÚ   CommDate
                    .Col = 3: strExamtype = Trim(.Text)                     '-- ±¸ºÐ       Examtype
                    .Col = 4: strBarCode = Trim(.Text)                      '-- ¹ÙÄÚµå     BarCode
                    
                          SQL = " Delete From PAT_RES  "
                    SQL = SQL & " Where CommDate = '" & strCommDate & "'"
                    SQL = SQL & "   and BarCode  = '" & strBarCode & "'"
                    SQL = SQL & "   and ExamType = '" & strExamtype & "'"
                    
                    Res = SendQuery(gLocal, SQL)
                End If
            Next
        End With
        
        cmdWorkSearch_Click
    End If
    
End Sub

Private Sub cmdDelGLU_Click()
    Dim intRow As Integer
    Dim strCommDate As String
    Dim strExamtype As String
    Dim strBarCode  As String
    
    If MsgBox("GLU ¿À´õ°¡ »èÁ¦µË´Ï´Ù. »èÁ¦ÇÏ½Ã°Ú½À´Ï±î?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
        With vasWorkPrint
            For intRow = 1 To .DataRowCnt
                .Row = intRow
                .Col = 1
                If .Value = "1" Then
                    .Col = 2: strCommDate = Format(Trim(.Text), "yyyymmdd") '-- ÀÇ·ÚÀÏÀÚ   CommDate
                    .Col = 3: strExamtype = Trim(.Text)                     '-- ±¸ºÐ       Examtype
                    .Col = 4: strBarCode = Trim(.Text)                      '-- ¹ÙÄÚµå     BarCode
                    
                          SQL = " Delete From PAT_RES  "
                    SQL = SQL & " Where CommDate = '" & strCommDate & "'"
                    SQL = SQL & "   and BarCode  = '" & strBarCode & "'"
                    SQL = SQL & "   and ExamType = '" & strExamtype & "'"
                    SQL = SQL & "   and ExamID = 'f'"
                    
                    Res = SendQuery(gLocal, SQL)
                End If
            Next
        End With
        
        cmdWorkSearch_Click
    End If
    
End Sub

Private Sub cmdLoad_Click()
    Dim PJumin As String
    Dim pGrid_Point As Integer
    Dim adoRS   As New ADODB.Recordset
    Dim strBarNum As String
'    Dim PJumin  As String
    
'''          SQL = " Select Company,HospCode,ChartNo,PatName,PatSex,PatAge,PatJumin,PatNo,CommDate,ExamNo,ExamID,ComExamID, "
'''    SQL = SQL & "        Specimen,Result,Reference,Remark,RsltDate,IOFlag,BarCode,ExamType "
'''    SQL = SQL & "   from pat_res "
'''    SQL = SQL & "  where commdate = '" & Format(cboWorkList.Text, "yyyymmdd") & "' "
''''    SQL = SQL & "    and (result = '' or result is null) "
'''    SQL = SQL & "  order by BarCode,ExamType "
    
    '-- 2013.06.10 ¼öÁ¤
          SQL = " Select Distinct Company,HospCode,ChartNo,PatName,PatSex,PatAge,PatJumin,PatNo,CommDate,BarCode,ExamType "
    SQL = SQL & "   from pat_res "
    SQL = SQL & "  where commdate = '" & Format(cboWorkList.Text, "yyyymmdd") & "' "
'    SQL = SQL & "    and (result = '' or result is null) "
    SQL = SQL & "  order by BarCode,ExamType "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open SQL, cn
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    With vasRePrint
        .MaxRows = adoRS.RecordCount
        Do While Not adoRS.EOF
            pGrid_Point = pGrid_Point + 1
'            .SetText 2, pGrid_Point, adoRS.Fields(0).Value & ""
'            .SetText 3, pGrid_Point, adoRS.Fields(1).Value & ""
'            .SetText 4, pGrid_Point, adoRS.Fields(2).Value & ""
'            .SetText 5, pGrid_Point, adoRS.Fields(3).Value & ""
'            .SetText 6, pGrid_Point, adoRS.Fields(4).Value & ""
'            .SetText 7, pGrid_Point, adoRS.Fields(5).Value & ""
'            .SetText 8, pGrid_Point, adoRS.Fields(6).Value & ""
'            .SetText 9, pGrid_Point, adoRS.Fields(7).Value & ""
'            .SetText 10, pGrid_Point, adoRS.Fields(8).Value & ""
'            .SetText 11, pGrid_Point, adoRS.Fields(9).Value & ""
'            .SetText 12, pGrid_Point, adoRS.Fields(10).Value & ""
'            .SetText 13, pGrid_Point, adoRS.Fields(11).Value & ""
'            .SetText 14, pGrid_Point, adoRS.Fields(12).Value & ""
'            .SetText 15, pGrid_Point, adoRS.Fields(13).Value & ""
'            .SetText 16, pGrid_Point, adoRS.Fields(14).Value & ""
'            .SetText 17, pGrid_Point, adoRS.Fields(15).Value & ""
'            .SetText 18, pGrid_Point, adoRS.Fields(16).Value & ""
'            .SetText 19, pGrid_Point, adoRS.Fields(17).Value & ""
            
            .SetText 1, pGrid_Point, "1"
            .SetText 2, pGrid_Point, Format(adoRS.Fields("CommDate").Value & "", "####-##-##")
            .SetText 3, pGrid_Point, Trim(adoRS.Fields("ExamType").Value & "")
            'strBarNum = Mid(Trim(adoRS.Fields("CommDate").Value & ""), 3) & Format(Trim(adoRS.Fields("ChartNo").Value & ""), "000000")
            strBarNum = Mid(adoRS.Fields("CommDate").Value & "", 5, 4) & Format(Trim(adoRS.Fields("ChartNo").Value & ""), "0000000000")
            .SetText 4, pGrid_Point, strBarNum
            .SetText 5, pGrid_Point, Trim(adoRS.Fields("ChartNo").Value & "")
            .SetText 6, pGrid_Point, Trim(adoRS.Fields("PatName").Value & "")
                        PJumin = Left(Trim(adoRS.Fields("PatJumin").Value & ""), 6) & Right(Trim(adoRS.Fields("PatJumin").Value & ""), 7)
                        Call CalAgeSex(PJumin, Format(Date, "yyyy/mm/dd"))
            .SetText 7, pGrid_Point, gPatGen.Sex
            .SetText 8, pGrid_Point, gPatGen.Age
            .SetText 9, pGrid_Point, "Print"
            
            adoRS.MoveNext
        Loop
        .RowHeight(-1) = 12
    End With
    
    adoRS.Close

End Sub

Private Sub WorkListLoad(ByVal strStartDate As String)
    Dim PJumin As String
    Dim pGrid_Point As Integer
    Dim adoRS   As New ADODB.Recordset
    
    'txtToday
    
          SQL = " Select distinct commdate "
    SQL = SQL & "   from pat_res "
    SQL = SQL & "  where commdate Between '" & Format(strStartDate, "yyyymmdd") & "' and '" & Format(Now, "yyyymmdd") & "' "
    'SQL = SQL & "    and (result = '' or result is null) "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open SQL, cn
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    cboWorkList.Clear
    Do While Not adoRS.EOF
        cboWorkList.AddItem Format(adoRS.Fields(0).Value & "", "####-##-##")
        adoRS.MoveNext
    Loop
    adoRS.Close

    
End Sub

Private Sub cmdPrint_Click()

    Dim intRow As Integer
    
    With vasPrint
        If .DataRowCnt <= 0 Then
            Exit Sub
        End If
        
        Timer1.Enabled = False
        
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = 1
            If .Value = "1" Then
                .Row = intRow
                .Col = 4
                If Trim(.Text) <> "" Then
                    If CP2140_Print(intRow, 0) Then
                        '-- Ãâ·Â ¼º°øÇÏ¸é
                        .Row = intRow
                        .Col = 1: .Value = "0"
                        .Col = 9: .Text = "Print"
                        
                        .Row = intRow
                        .Row2 = intRow
                        .Col = 1
                        .Col2 = .MaxCols
                        .BlockMode = True
                        .BackColor = vbCyan
                        .BlockMode = False
                    Else
                        '-- Ãâ·Â ½ÇÆÐÇÏ¸é
                        .Row = intRow
                        .Col = 9: .Text = "Error"
                    
                    End If
                End If
                .Row = intRow
                .Col = 1
                .Value = "0"
            End If
        Next
    
        lblTimer.Caption = gBar_Parm.WaitSec
        lngRefresh = gBar_Parm.WaitSec
        
        Timer1.Interval = 1000
        Timer1.Enabled = True
    
    End With
    
End Sub

Private Function ArgoxIntermecString(ByVal pString As String) As String
    ArgoxIntermecString = Chr(34) & pString & Chr(34) & vbLf
End Function

Private Function ArgoxPrint(ByVal pIntRow As Integer, ByVal Idx As Integer)
    Dim strTestNm1  As String
    Dim strTestNm2  As String
    Dim strQValue1  As String
    Dim strQValue2  As String
    Dim strqValue   As String
    Dim strFont     As String
    Dim strFrozenFg As String
    Dim strPosition As String
    Dim strBarFor   As String
    Dim strDiv      As String
    Dim strTmpSpcno As String
    
    Dim ii          As Integer
    
    Dim strColDt    As String
    Dim strPart     As String
    Dim strBarNum   As String
    Dim strChart    As String
    Dim strPatNm    As String
    Dim strPatSex   As String
    Dim strPatAge   As String
    Dim strTestNm   As String
    
'-- ½Ã¸®¾ó Æ÷Æ® ¿ÀÇÂ
On Error GoTo Errors
    
    If Idx = 0 Then
        With vasPrint
            .Row = pIntRow
            .Col = 2: strColDt = Trim(.Text)
            .Col = 3: strPart = Trim(.Text)
            .Col = 4: strBarNum = Trim(.Text)
            .Col = 5: strChart = Trim(.Text)
            .Col = 6: strPatNm = Trim(.Text)
            .Col = 7: strPatSex = Trim(.Text)
            .Col = 8: strPatAge = Trim(.Text)
        End With
    Else
        With vasRePrint
            .Row = pIntRow
            .Col = 2: strColDt = Trim(.Text)
            .Col = 3: strPart = Trim(.Text)
            .Col = 4: strBarNum = Trim(.Text)
            .Col = 5: strChart = Trim(.Text)
            .Col = 6: strPatNm = Trim(.Text)
            .Col = 7: strPatSex = Trim(.Text)
            .Col = 8: strPatAge = Trim(.Text)
        End With
    End If
    
    Select Case strPart
'        Case "C": strPart = "Chemistry"
'        Case "H": strPart = "Hematology"
'        Case "I": strPart = "Immunology"
'        Case "U": strPart = "Urine"
        Case "C": strPart = "Chemi"
        Case "H": strPart = "Hemato"
        Case "I": strPart = "Immuno"
        Case "U": strPart = "Urine"
    End Select
    
    If MSComm1.PortOpen = False Then
        MSComm1.CommPort = gSetup.gPort
        MSComm1.RTSEnable = gSetup.gRTSEnable
        MSComm1.DTREnable = gSetup.gDTREnable
        MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
        MSComm1.PortOpen = True
    End If

    With MSComm1
        .Output = "N" + vbCrLf                                  'Å¬¸®¾î ¹öÆÛ
        '.Output = "KI82" & vbCrLf                             'Åõ°úÇü¼¾¼­¸¸»ç¿ë½Ã k182 // ±âº»¼¾¼­(¹Ý»çÇü) k180
        .Output = "JF" + vbCrLf                                 '»ªÇÇµå
        .Output = "OD" + vbCrLf                                 '°¨¿­ »ç¿ë ¿É¼Ç
        .Output = "D10" & vbCrLf                                'Density ±âº»°ªÀº 10 °ªÀÌ ³ôÀ»¼ö·Ï ÁøÇÏ°Ô
        .Output = "S2" + vbCrLf                                 '½ºÇÇµå 2 ÀûÁ¤°ª
        .Output = "Q" & CStr(gBar_Parm.LabelHeight) & "," & CStr(gBar_Parm.LabelGap) & vbCrLf    'Q°ª 3.5*0.8*100=280[¶óº§ÀÇ ±æÀÌ], 0.8*0.3cm(°Ø)=24
        .Output = "q" & CStr(gBar_Parm.LabelWidth) & vbCrLf                       '5.4*0.8*100=432[¶óº§ÀÇ Width]
        .Output = "ZB" + vbCrLf                                 'ÇÁ¸°Æ® ÀÌ¹ÌÁö
        
        '±âÅ¸ µ¥ÀÌÅ¸ (ÀÇ·ÚÀÏ,           ¹ÙÄÚµå¹øÈ£
        '             °Ë»çÇ×¸ñ
        '             ¼öÁøÀÚ¸í,         ¼ºº°,³ªÀÌ
        '+ °Ë»çÇ×¸ñ
        
        For ii = 1 To 7
            'ÇÑ±ÛÀÎ°æ¿ì '0'À» '9'·Î º¯°æ
            'strFont = mudtBarData(ii).FontX
            'If strFont = "0" Then strFont = "9"
            'strPosition = "A" & mudtBarData(ii).PosX & "," & mudtBarData(ii).PosY & ",0," & strFont & ",1,1,N,"
            Select Case ii
                Case 1: '-- ÀÇ·ÚÀÏ
                        strFont = "9"
                        strPosition = "A" & gBar_Parm.ColDateX & "," & gBar_Parm.ColDateY & ",0," & strFont & ",1,1,N,"
                        'strPosition = "A" & "20" & "," & "30" & ",0," & strFont & ",1,1,N,"
                        .Output = strPosition & ArgoxIntermecString(strColDt)                                          '
'                Case 2: '-- ¹ÙÄÚµå ¹øÈ£
''                        strFont = "2"
''                        'strPosition = "A" & gBar_Parm.BarNumX & "," & gBar_Parm.BarNumY & ",0," & strFont & ",1,1,N,"
''                        strPosition = "A" & "50" & "," & "20" & ",0," & strFont & ",1,1,N,"
''                        .Output = strPosition & ArgoxIntermecString(strBarNum)
                Case 3: '-- ¹ÙÄÚµå ¶óº§
                        'strFont = "3"
                        strBarFor = "3"
                        strPosition = "B" & gBar_Parm.BarCodeX & "," & gBar_Parm.BarCodeY & ",0," & strBarFor & ",2,3," & gBar_Parm.BarCodeH & ",B,"    '-- ¹ÙÄÚµå¹øÈ£ Ãâ·Â
                        'strPosition = "B" & gBar_Parm.BarCodeX & "," & gBar_Parm.BarCodeY & ",0," & strBarFor & ",2,4," & gBar_Parm.BarCodeH & ",N,"    '-- ¹ÙÄÚµå¹øÈ£ ¹ÌÃâ·Â
                        'strPosition = "B" & "30" & "," & "70" & ",0," & strBarFor & ",2,4," & "90" & ",B,"'-- ¹ÙÄÚµå¹øÈ£ Ãâ·Â
                        .Output = strPosition & ArgoxIntermecString(strBarNum)
                Case 4: '-- °Ë»çÇ×¸ñ
                        strTestNm = strPart
                        strFont = "9"
                        strPosition = "A" & gBar_Parm.TestNmX & "," & gBar_Parm.TestNmY & ",0," & strFont & ",1,1,N,"
                        'strPosition = "A" & "30" & "," & gBar_Parm.TestNmY & ",0," & strFont & ",1,1,N,"
                        .Output = strPosition & ArgoxIntermecString(strTestNm)
                Case 5: '-- ¼öÁøÀÚ¸í
                        strFont = "9"
                        strPosition = "A" & gBar_Parm.PatNmX & "," & gBar_Parm.PatNmY & ",0," & strFont & ",1,1,N,"
                        'strPosition = "A" & "30" & "," & "240" & ",0," & strFont & ",1,1,N,"
                        .Output = strPosition & ArgoxIntermecString(strPatNm)
'                            .Output = "A10,140,0,9,1,1,N," & ArgoxIntermecString("Á¤ÀÏÈ¯")
                Case 6: '-- ¼ºº°
                        strFont = "9"
                        strPosition = "A" & gBar_Parm.PatSexX & "," & gBar_Parm.PatSexY & ",0," & strFont & ",1,1,N,"
                        'strPosition = "A" & gBar_Parm.PatSexX & "," & "240" & ",0," & strFont & ",1,1,N,"
                        .Output = strPosition & ArgoxIntermecString(strPatSex & "/" & strPatAge)
'                Case 7: '-- ³ªÀÌ
                        'strFont = "9"
                        ''strPosition = "A" & gBar_Parm.PatAgeX & "," & gBar_Parm.PatAgeY & ",0," & strFont & ",1,1,N,"
                        'strPosition = "A" & "160" & "," & "240" & ",0," & strFont & ",1,1,N,"
                        '.Output = strPosition & ArgoxIntermecString(strPatAge)
            End Select
        Next
        
        .Output = "P1," & txtCnt.Text & vbCrLf
'        .Output = "P" & "1,1" & vbCrLf
        
        ' -- ½Ã¸®¾óÆ÷Æ® Close
        If .PortOpen Then .PortOpen = False
    
    End With

Errors:

End Function


Private Function CP2140_Print(ByVal intRow As Integer, ByVal Idx As Integer) As Boolean

    On Error GoTo ErrPrint
    
    CP2140_Print = False
    
    Call ArgoxPrint(intRow, Idx)
    
    'MSComm1.Output = strBarNo
    
    CP2140_Print = True
    
Exit Function

ErrPrint:
    CP2140_Print = False

End Function

Private Sub cmdPrtSetup_Click()

    frmConfig.Show 1
    
End Sub

Private Sub cmdRePrint_Click()
    Dim intRow As Integer
    
    With vasRePrint
        If .DataRowCnt <= 0 Then
            Exit Sub
        End If
        
        Timer1.Enabled = False
        
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = 1
            If .Value = "1" Then
                .Row = intRow
                .Col = 4
                
                If CP2140_Print(intRow, 1) Then
                    '-- Ãâ·Â ¼º°øÇÏ¸é
                    .Row = intRow
                    .Col = 1: .Value = "0"
                    .Col = 9: .Text = "Print"
                    
                    .Row = intRow
                    .Row2 = intRow
                    .Col = 1
                    .Col2 = .MaxCols
                    .BlockMode = True
                    .BackColor = vbCyan
                    .BlockMode = False
                Else
                    '-- Ãâ·Â ½ÇÆÐÇÏ¸é
                    .Row = intRow
                    .Col = 9: .Text = "Error"
                
                End If
            End If
        Next
        lblTimer.Caption = gBar_Parm.WaitSec
        lngRefresh = gBar_Parm.WaitSec
        
        Timer1.Interval = 1000
    
    End With

End Sub

Private Sub cmdReSearch_Click()

    Call WorkListLoad(dtpBeginDate.Value)

End Sub

Private Sub cmdSearch_Click()
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
    Dim PName As String
    Dim PJumin As String
    Dim PID As String
    Dim PExamCode As String
    Dim PReceDate As String
    Dim PAge As String
    Dim pSex As String
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
    
    Screen.MousePointer = 11
    
    ClearSpread vasWork

    varXML = f_subSet_XMLWorkList(Format(dtpStartDt.Value, "yyyymmdd"), Format(dtpStopDt.Value, "yyyymmdd"))
    
    If blnSameRecord = False Then
        'MsgBox "°Ë»ç ´ë»óÀÚ°¡ ¾ø½À´Ï´Ù.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    End If
    
    If UBound(varXML) < 1 Then
        'MsgBox "°Ë»ç ´ë»óÀÚ°¡ ¾ø½À´Ï´Ù.", vbOKOnly + vbInformation, App.Title
        Exit Sub
    Else
        strBarNo = ""

        With vasPrint
            '.Visible = False
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
                    
                    SQL = "select equipno, equipcode, examname, examtype from equipexam where examcode = '" & XMLInData.ExamID & "' "
                    Res = db_select_Col(gLocal, SQL)
    '                Debug.Print XMLInData.ExamID
                    If Res > 0 Then
                        PEquipno = gReadBuf(0)
                        PEquipCode = gReadBuf(1)
                        PExamname = gReadBuf(2)
                                        
                        If strBarNo <> XMLInData.ChartNo Or pEqipType <> gReadBuf(3) Then
                            pEqipType = gReadBuf(3)
                            
                            pGrid_Point = SeqSearch_New(vasPrint, XMLInData.ChartNo, pEqipType, 5)
        
                            If pGrid_Point = 0 Then
                                pGrid_Point = SeqNullSearch(vasPrint, XMLInData.ChartNo, 5)
                                If pGrid_Point = 0 Then .MaxRows = .MaxRows + 1: pGrid_Point = .MaxRows
                                .RowHeight(-1) = 12
                            End If
                            
                            If chkAuto.Value = "1" Then
                                .SetText 1, pGrid_Point, "1"
                            Else
                                .SetText 1, pGrid_Point, "0"
                            End If
                            
                            .SetText 2, pGrid_Point, Format(XMLInData.CommDate, "####-##-##")
                            .SetText 3, pGrid_Point, pEqipType
                            strBarNum = Mid(XMLInData.CommDate, 5, 4) & Format(XMLInData.ChartNo, "0000000000")
                            'strBarNum = Format$(XMLInData.ChartNo, String$(SPCLEN, "#"))
                            
                            .SetText 4, pGrid_Point, strBarNum
                            .SetText 5, pGrid_Point, XMLInData.ChartNo
                            .SetText 6, pGrid_Point, XMLInData.PatName
                                        PJumin = Left(XMLInData.PatJumin, 6) & Right(XMLInData.PatJumin, 7)
                                        Call CalAgeSex(PJumin, Format(Date, "yyyy/mm/dd"))
                            .SetText 7, pGrid_Point, gPatGen.Sex
                            .SetText 8, pGrid_Point, gPatGen.Age
                            .SetText 9, pGrid_Point, "Order"
    
                        End If
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
                                SQL = SQL & "'" & XMLInData.Remark & "',"
                                SQL = SQL & "'" & XMLInData.RsltDate & "',"
                                SQL = SQL & "'" & XMLInData.IOFlag & "',"
                                SQL = SQL & "'" & strBarNum & "',"
                                SQL = SQL & "'" & pEqipType & "')"
                                
                                Res = SendQuery(gLocal, SQL)
                                
                                If Res = -1 Then
                                    SaveQuery SQL
                                End If
                            
                            '-- ¼ÓµµÇâ»óÀ» À§ÇØ Äõ¸®¹® Áö¿ì±â
                            Else
                                      SQL = " Update pat_res Set "
                                SQL = SQL & " PatName = '" & XMLInData.PatName & "', "
                                SQL = SQL & " PatSex  = '" & gPatGen.Sex & "' "
                                'SQL = SQL & " ExamNo = '" & XMLInData.ExamNo & "', "
                                'SQL = SQL & " PatNo = '" & XMLInData.PatNo & "',"
                                SQL = SQL & " Where ChartNo  = '" & XMLInData.ChartNo & "' "
                                SQL = SQL & "   and ExamID   = '" & XMLInData.ExamID & "' "
                                SQL = SQL & "   and CommDate = '" & XMLInData.CommDate & "'"
                                SQL = SQL & "   and BarCode  = '" & strBarNum & "'"
                                SQL = SQL & "   and ExamType = '" & pEqipType & "'"
                                
                                Res = SendQuery(gLocal, SQL)
                            End If
                            
                                                    
                            strBarNo = XMLInData.ChartNo
                        'End If
                    End If
                Else
                    'XMLInData.ComExamID = ""
                End If
                XMLInData.ComExamID = ""
            Next
            
            If chkAuto.Value = "1" Then
                Call cmdPrint_Click
            End If
        End With
    End If
    
    'Kill "C:\UBCare\SINAI\IF\ExamIF_In.xml"
    
    strSrcfile = "C:\UBCare\SINAI\IF\ExamIF_In.xml"
    strDestFile = App.Path & "\Log\" & "ExamIF_In_" & Format(Now, "yymmddhhmmss") & ".xml"

    FileCopy strSrcfile, strDestFile
    Kill strSrcfile

    Screen.MousePointer = 0
    'Exit Sub
    
End Sub

Private Sub cmdSetup_Click()
'    frmEquipExam.SSPanel1.Caption = "  PhD Àåºñ ÄÚµå ¼³Á¤"
    frmEquipExam.Show 1
    GetExamCode
End Sub

Private Sub cmdWorkDel_Click()
    Dim SQL As String
    Dim intRow As Integer
    Dim intNum As Integer
    
    For intRow = 1 To vasList.DataRowCnt
        vasList.Row = intRow
        vasList.Col = 1
        
        If vasList.Value = "1" Then
                  SQL = " Delete From pat_res  "
            vasList.Col = 4
            SQL = SQL & " Where ChartNo  = '" & Trim(vasList.Text) & "' "
            vasList.Col = 12
            SQL = SQL & "   and ExamID   = '" & Trim(vasList.Text) & "' "
            vasList.Col = 10
            SQL = SQL & "   and CommDate = '" & Trim(vasList.Text) & "'"
            Res = SendQuery(gLocal, SQL)
            'Call vasList.DeleteRows(intRow, intRow)
        End If
    Next
    
    intNum = 0
    For intRow = 1 To vasList.DataRowCnt
        vasList.Row = intRow
        vasList.Col = 1
        
        If vasList.Value = "0" Then
            intNum = intNum + 1
                  SQL = " Update pat_res Set "
            SQL = SQL & " RSLTDATE = '" & Format(txtToday.Text, "yyyymmdd") & "', "
            SQL = SQL & " REMARK  = '" & intNum & "' "
            vasList.Col = 4
            SQL = SQL & " Where ChartNo  = '" & Trim(vasList.Text) & "' "
            vasList.Col = 12
            SQL = SQL & "   and ExamID   = '" & Trim(vasList.Text) & "' "
            vasList.Col = 10
            SQL = SQL & "   and CommDate = '" & Trim(vasList.Text) & "'"
            Res = SendQuery(gLocal, SQL)
        End If
    Next
    
    Call cmdLoad_Click
    
    
End Sub

Private Sub cmdWorkList_Click()

    Dim varTmp  As Variant
    Dim intRow1 As Integer, intRow2 As Integer
    Dim intIdx  As Integer
    Dim Rev     As Long
    Dim Test_Cd() As String, strPid()   As String, strPnm() As String
    Dim blnFlag As Boolean
    Dim strBarNo    As String, strSPid  As String, strSPnm   As String, strChartNo As String, strSex As String
    Dim strWDate As String
    Dim strEqpCd    As String
    Dim tmpDate     As String
    Dim strORDT, strORQN, strPANM, strPAID, strOIFL, strSENO, strSEXS, strAGES, strNWNO, strORCD, strSQNO As String
    Dim strData(20) As String
    
    blnFlag = False
    
    With vasWork
        For intRow1 = 1 To .MaxRows
            .GetText 1, intRow1, varTmp
            If Trim$(varTmp) = "1" Then
                .GetText 2, intRow1, varTmp:    strData(2) = Trim$(varTmp)
                .GetText 3, intRow1, varTmp:    strData(3) = Trim$(varTmp)
                .GetText 4, intRow1, varTmp:    strData(4) = Trim$(varTmp)
                .GetText 5, intRow1, varTmp:    strData(5) = Trim$(varTmp)
                .GetText 6, intRow1, varTmp:    strData(6) = Trim$(varTmp)
                .GetText 7, intRow1, varTmp:    strData(7) = Trim$(varTmp)
                .GetText 8, intRow1, varTmp:    strData(8) = Trim$(varTmp)
                .GetText 9, intRow1, varTmp:    strData(9) = Trim$(varTmp)
                .GetText 10, intRow1, varTmp:   strData(10) = Trim$(varTmp)
                .GetText 11, intRow1, varTmp:   strData(11) = Trim$(varTmp)
                .GetText 12, intRow1, varTmp:   strData(12) = Trim$(varTmp)
                .GetText 13, intRow1, varTmp:   strData(13) = Trim$(varTmp)
                .GetText 14, intRow1, varTmp:   strData(14) = Trim$(varTmp)
                .GetText 15, intRow1, varTmp:   strData(15) = Trim$(varTmp)
                .GetText 16, intRow1, varTmp:   strData(16) = Trim$(varTmp)
                .GetText 17, intRow1, varTmp:   strData(17) = Trim$(varTmp)
                .GetText 18, intRow1, varTmp:   strData(18) = Trim$(varTmp)
                .GetText 19, intRow1, varTmp:   strData(19) = Trim$(varTmp)
                
                .Row = intRow1: .Col = 1: .ForeColor = vbRed
                                .Col = 2: .ForeColor = vbRed
                                .Col = 3: .ForeColor = vbRed
                                .Col = 4: .ForeColor = vbRed
                                .Col = 5: .ForeColor = vbRed
                                .Col = 6: .ForeColor = vbRed
                                .Col = 7: .ForeColor = vbRed
                                .Col = 8: .ForeColor = vbRed
                                .Col = 9: .ForeColor = vbRed
                                .Col = 10: .ForeColor = vbRed
                                
                intRow2 = SeqSearch(vasList, strData(4), 4)
                If intRow2 < 1 Then
                    intRow2 = SeqNullSearch(vasList, strData(4), 4)
                    If intRow2 < 1 Then
                        vasList.MaxRows = vasList.MaxRows + 1
                        vasList.RowHeight(vasList.MaxRows) = 12
                        intRow2 = vasList.MaxRows
                    End If

                    'blnFlag = False
                    
                    'If blnFlag = True Then
                        vasList.SetText 2, intRow2, strData(2)
                        vasList.SetText 3, intRow2, strData(3)
                        vasList.SetText 4, intRow2, strData(4)
                        vasList.SetText 5, intRow2, strData(5)
                        vasList.SetText 6, intRow2, strData(6)
                        vasList.SetText 7, intRow2, strData(7)
                        vasList.SetText 8, intRow2, strData(8)
                        vasList.SetText 9, intRow2, strData(9)
                        vasList.SetText 10, intRow2, strData(10)
                        vasList.SetText 11, intRow2, strData(11)
                        vasList.SetText 12, intRow2, strData(12)
                        vasList.SetText 13, intRow2, strData(13)
                        vasList.SetText 14, intRow2, strData(14)
                        vasList.SetText 15, intRow2, strData(15)
                        vasList.SetText 16, intRow2, strData(16)
                        vasList.SetText 17, intRow2, strData(17)
                        vasList.SetText 18, intRow2, strData(18)
                        vasList.SetText 19, intRow2, strData(19)
                        
                        vasList.Row = intRow2:
                        vasList.Col = 2:
                        vasList.ForeColor = vbRed
                    'Else
                    '    vasList.MaxRows = vasList.MaxRows - 1
                    'End If
                End If
                
                .SetText 1, intRow1, ""
            End If
        Next
    End With
                
End Sub


Private Sub cmdWorkPrint_Click()

    With vasWorkPrint
        If .DataRowCnt > 0 Then
            .PrintOrientation = PrintOrientationPortrait
            '.PrintOrientation = PrintOrientationLandscape
            .Action = ActionPrint
            MsgBox dtpSearch.Value & " ÀÏÀÚ [" & cboPart.Text & "] ÀÇ ¿öÅ©¸®½ºÆ®°¡ Ãâ·ÂµÇ¾ú½À´Ï´Ù..       " & vbCrLf & vbCrLf & "´ÙÀ½ ÀÛ¾÷À» ÁøÇàÇÏ½Ê½Ã¿ä..", vbInformation + vbOKOnly, App.Title
        End If
    End With
    
End Sub

Private Sub cmdWorkSave_Click()
    Dim SQL As String
    Dim intRow As Integer
    
    For intRow = 1 To vasList.DataRowCnt
        vasList.Row = intRow
        vasList.Col = 1
        
        If vasList.Value = "1" Then
                  SQL = " Update pat_res Set "
            SQL = SQL & " RSLTDATE = '" & Format(txtToday.Text, "yyyymmdd") & "', "
            SQL = SQL & " REMARK  = '" & intRow & "' "
            vasList.Col = 4
            SQL = SQL & " Where ChartNo  = '" & Trim(vasList.Text) & "' "
            vasList.Col = 12
            SQL = SQL & "   and ExamID   = '" & Trim(vasList.Text) & "' "
            vasList.Col = 10
            SQL = SQL & "   and CommDate = '" & Trim(vasList.Text) & "'"
            Res = SendQuery(gLocal, SQL)
        End If
    Next
    
End Sub


Private Sub cmdWorkSearch_Click()
    Dim PJumin As String
    Dim pGrid_Point As Integer
    Dim adoRS   As New ADODB.Recordset
    Dim strBarNum As String
    Dim strExamName As String
    Dim varTmp As Variant
    
    ClearSpread vasWorkPrint

    '-- 2013.06.12 ¼öÁ¤
          SQL = " Select a.Company,a.HospCode,a.ChartNo,a.PatName,a.PatSex,a.PatAge,a.PatJumin,a.PatNo,a.CommDate,a.BarCode,a.ExamType,b.examname"
    SQL = SQL & "   from pat_res a, equipexam b"
    SQL = SQL & "  where commdate = '" & Format(dtpSearch.Value, "yyyymmdd") & "' "
    SQL = SQL & "    and a.examtype = b.examtype"
    SQL = SQL & "    and a.ExamID =b.examcode"
  
    If Mid(cboPart, 1, 1) <> "A" Then
        SQL = SQL & "   and a.Examtype = '" & Mid(cboPart, 1, 1) & "'"
    End If
    
    SQL = SQL & "  order by a.PatName,a.BarCode, a.ExamID "
    
    adoRS.CursorLocation = adUseClient
    adoRS.Open SQL, cn
    If adoRS.RecordCount > 0 Then adoRS.MoveFirst
    With vasWorkPrint
        '.MaxRows = adoRS.RecordCount
        Do While Not adoRS.EOF
            If strBarNum <> Mid(adoRS.Fields("CommDate").Value & "", 5, 4) & Format(Trim(adoRS.Fields("ChartNo").Value & ""), "0000000000") Then
                pGrid_Point = pGrid_Point + 1
                .MaxRows = pGrid_Point
                .SetText 1, pGrid_Point, "0"
                .SetText 2, pGrid_Point, Format(adoRS.Fields("CommDate").Value & "", "####-##-##")
                .SetText 3, pGrid_Point, Trim(adoRS.Fields("ExamType").Value & "")
                'strBarNum = Mid(Trim(adoRS.Fields("CommDate").Value & ""), 3) & Format(Trim(adoRS.Fields("ChartNo").Value & ""), "000000")
                strBarNum = Mid(adoRS.Fields("CommDate").Value & "", 5, 4) & Format(Trim(adoRS.Fields("ChartNo").Value & ""), "0000000000")
                .SetText 4, pGrid_Point, strBarNum
                .SetText 5, pGrid_Point, Trim(adoRS.Fields("ChartNo").Value & "")
                .SetText 6, pGrid_Point, Trim(adoRS.Fields("PatName").Value & "")
                            PJumin = Left(Trim(adoRS.Fields("PatJumin").Value & ""), 6) & Right(Trim(adoRS.Fields("PatJumin").Value & ""), 7)
                            Call CalAgeSex(PJumin, Format(Date, "yyyy/mm/dd"))
                .SetText 7, pGrid_Point, gPatGen.Sex
                .SetText 8, pGrid_Point, gPatGen.Age
                .SetText 9, pGrid_Point, "Print"
            End If
            
            .GetText 10, pGrid_Point, varTmp
            .SetText 10, pGrid_Point, varTmp & IIf(varTmp <> "", ",", "") & Trim(adoRS.Fields("ExamName").Value & "")
            
            adoRS.MoveNext
        Loop
        .RowHeight(-1) = 12
    End With
    
    adoRS.Close
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
    
    strBufferData = ""
    strBufferData = ""
    strBufferData = strBufferData & "1H|\^&|||PhD_EIA^2^0^0^14|||||||P|LIS2-A2|20120209135511" & vbCr
    strBufferData = strBufferData & "2C" & vbCr
    strBufferData = strBufferData & "2P|1||||^^^^" & vbCr
    strBufferData = strBufferData & "BB" & vbCr
    strBufferData = strBufferData & "3O|1|51062^ccp.02.09.2012^1||^^^Anti-ccp|R||||||||||||||||||||F" & vbCr
    strBufferData = strBufferData & "CF" & vbCr
    strBufferData = strBufferData & "4R|1|^^^Anti-ccp^QUANT|0.17|U/ml||||F" & vbCr
    strBufferData = strBufferData & "14" & vbCr
    strBufferData = strBufferData & "5R|2|^^^Anti-ccp^QUAL^^^F|neg|||||F" & vbCr
    strBufferData = strBufferData & "37" & vbCr
    strBufferData = strBufferData & "6L|1|N09" & vbCr
    strBufferData = strBufferData & "" & vbCr

    Call MSComm1_OnComm
    
End Sub


Private Sub dtpBeginDate_Change()

    Call cmdReSearch_Click
    
End Sub

Private Sub Form_Activate()
    txtMsg.Text = ""
End Sub

Private Sub cmdRun()
    
    If Not MSComm1.PortOpen Then MSComm1.PortOpen = True
    
    If MSComm1.PortOpen Then
        lblStatus.Caption = "¿¬°á µÇ¾ú½À´Ï´Ù."
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
'        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    Else
        lblStatus.Caption = "¿¬°á µÇÁö ¾Ê¾Ò½À´Ï´Ù."
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
'        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
'        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    End If
        
End Sub

Private Sub Form_Load()
    Dim sDate As String
    '1. È­¸é ¹× º¯¼ö ÃÊ±âÈ­
    '2. µ¥ÀÌÅ¸º£ÀÌ½º¿¡ Connect ÇÏ±â - Local - Server
    '3. Ini ³»¿ë ºÒ·¯¿À±â    GetSetup
    '4. Comport Open

'    Me.Left = 0
'    Me.Top = 0
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    
    stInterface.Tab = 0
    txtMsg.Text = ""

    ClearSpread vasPrint
    ClearSpread vasRePrint
    ClearSpread vasID
    ClearSpread vasWork
    ClearSpread vasList
    'vasActiveCell vasID, 1, colPID
    
'    vasRes.OperationMode = 0
'    ClearSpread vasRes, 1, 1
'    vasRes.MaxRows = 0
    
    GetSetup    'ini¿¡¼­ DBÁ¤º¸ ºÒ·¯¿À±â
        
'    If Not Connect_Server Then
'        MsgBox "¼­¹ö¿¡ ¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
'        'Exit Sub
'    End If

    If Not Connect_Local Then
        MsgBox "¿¬°áµÇÁö ¾Ê¾Ò½À´Ï´Ù."
        'Exit Sub
    End If
    
    If gAutoSend = 1 Then
        optAuto.Value = True
    Else
        optManual = True
    End If
    
    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit
    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    Call cmdRun
    'Me.txtUID = gExamUID
    
    raw_data = ""
    
    txtToday = Format(CDate(GetDateFull), "yyyy/mm/dd")
    
    '====================·ÎÄÃ DBÁö¿ì±â - 30ÀÏ º¸°ü======================
    sDate = Format(DateAdd("y", CDate(txtToday.Text), -30), "yyyymmdd")
    
    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    SendQuery gLocal, SQL
    
    SQL = "Delete from Worklist where examdate < '" & sDate & "' "
    SendQuery gLocal, SQL
    
    '===================================================================
    
    '°Ë»çÄÚµå °¡Á®¿À±â
    GetExamCode
        
    'MultiSelect Mode
'    vasRes.OperationMode = 1
    
'    SQL = " Alter table pat_res Alter column recedate text(20) "
'    res = SendQuery(gLocal, SQL)
    
'    dtpSDate.Text = Format(DateAdd("y", CDate(GetDateFull), -3), "yyyy/mm/dd")
'    dtpEDate.Text = Format(CDate(GetDateFull), "YYYY/MM/DD")
    
    ClearSpread vasList
    ClearSpread vasWorkPrint
    
    ChkAll1.Value = 0
        
    '==============================
    intPhase = 1
    strState = ""
    intBufCnt = 0
    blnIsETB = False
    intSndPhase = 0
    intFrameNo = 1
    '==============================
    
    'dtpStartDt.Value = Now - 3
    'dtpStartDt.Value = Format(Now, "yyyy") & "-01-01"
    
    dtpStartDt.Value = DateAdd("D", -30, Now)
    dtpStopDt.Value = Now
    
    dtpBeginDate.Value = Now - 3

    dtpSearch.Value = Now
    
    cboPart.Clear
    'cboPart.AddItem "All : ÀüÃ¼"
    cboPart.AddItem "Chemistry"
    cboPart.AddItem "Hematology"
    cboPart.AddItem "Immunology"
    cboPart.AddItem "Urine"
    cboPart.ListIndex = 0
    
    blnSameRecord = False
    
    Call WorkListLoad(dtpBeginDate.Value)
    
    lblTimer.Caption = gBar_Parm.WaitSec
    lngRefresh = gBar_Parm.WaitSec
    
    Timer1.Interval = 1000
    Timer1.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False

    'Call WritePrivateProfileString("EquipConf", "StartSeq", txtSNo.Text, App.Path & "\cp2140.ini")
    DisConnect_Local
    DisConnect_Server
    'DisConnect_Server1
    
    Unload Me
End Sub

Sub GetExamCode()
'°Ë»çÄÚµå¸¦ array¿¡ ÀúÀå
    Dim i As Integer
    Dim j As Integer
    
    gAllExam = ""
    
    ClearSpread vasTemp
    
    SQL = "Select EquipCode, ExamCode, ExamName From EquipExam where equipno = '" & gEquip & "' " & vbCrLf & _
          " Order by EquipCode"
          
    Res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If Res > 0 Then
        ReDim gArr_ExamCode(1 To vasTemp.DataRowCnt, 1 To 3)
    Else
        SaveQuery SQL
    End If
    
    For i = 1 To vasTemp.DataRowCnt
        gArr_ExamCode(i, 1) = i
        
        For j = 1 To 2
            gArr_ExamCode(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
        
        If gAllExam = "" Then
            gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ", '" & Trim(GetText(vasTemp, i, 2)) & "'"
        End If
    Next i
    
End Sub
'
'Private Sub monvCal_DateClick(ByVal DateClicked As Date)
'    If iIndex = 0 Then
'        dtpSDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
'    Else
'        dtpEDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
'    End If
'    monvCal.Visible = False
'End Sub



Private Sub lblTimer_DblClick()


    If Timer1.Enabled = True Then
        Timer1.Enabled = False
        lblTimer.BackColor = &HE0E0E0
    Else
        Timer1.Enabled = True
        lblTimer.BackColor = vbWhite
    End If
    
End Sub

Private Sub MSComm1_OnComm()

        
Dim strERMsg As String
Dim strEVMsg As String

    Select Case MSComm1.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long

            Buffer = MSComm1.Input
            'Buffer = strBufferData
            Save_Raw_Data "[Rx]" & Buffer
            lngBufLen = Len(Buffer)
            
            Debug.Print Buffer
            
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
                                'If strState = "Q" Then Call SendOrder
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
                                'Call EditRcvData
'                                Call Emerald(strRecvData)
'                                If strState = "Q" Then
'                                    intSndPhase = 1
'                                    intFrameNo = 1
'                                    MSComm1.Output = ENQ
'                                    Save_Raw_Data "[Tx]" & ENQ
'                                End If
                                intPhase = 1
                        End Select
                End Select
            Next i
        Case comEvSend
        
            'imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
'            If tmrSend.Enabled = False Then
'                tmrSend.Enabled = True
'            Else
'                tmrSend.Enabled = False
'                tmrSend.Enabled = True
'            End If
        Case comEvCTS
            strEVMsg = " CTS(Clear to Send) º¯°æ °¨Áö"
        Case comEvDSR
            strEVMsg = " DSR(Data Set Read) º¯°æ °¨Áö"
        Case comEvCD
            strEVMsg = " CD(Carrier Detect) º¯°æ °¨Áö"
        Case comEvRing
            strEVMsg = " ÀüÈ­ º§ÀÌ ¿ï¸®´Â Áß"
        Case comEvEOF
            strEVMsg = " EOF(End Of File) °¨Áö"

        ' ¿À·ù ¸Þ½ÃÁö
        Case comBreak
            strERMsg = " Áß´Ü ½ÅÈ£ ¼ö½Å"
        Case comCDTO
            strERMsg = " ¹Ý¼ÛÆÄ °ËÃâ ½Ã°£ ÃÊ°ú"
        Case comCTSTO
            strERMsg = " CTS(Clear to Send) ½Ã°£ ÃÊ°ú"
        Case comDCB
            strERMsg = " Æ÷Æ®¿¡ ´ëÇÑ ÀåÄ¡ Á¦¾î ºí·Ï(DCB) °Ë»ö Áß ¿¹±âÄ¡ ¸øÇÑ ¿À·ù"
        Case comDSRTO
            strERMsg = " DSR(Data Set Read) ½Ã°£ ÃÊ°ú"
        Case comFrame
            strERMsg = " ÇÁ·¹ÀÌ¹Ö ¿À·ù"
        Case comOverrun
            strERMsg = " ÆÐ¸®Æ¼ ¿À·ù"
        Case comRxOver
            strERMsg = " ¼ö½Å ¹öÆÛ ÃÊ°ú"
        Case comRxParity
            strERMsg = " ÆÐ¸®Æ¼ ¿À·ù"
        Case comTxFull
            strERMsg = " Àü¼Û ¹öÆÛ¿¡ ¿©À¯°¡ ¾øÀ½"
        Case Else
            strERMsg = " ¾Ë ¼ö ¾ø´Â ¿À·ù ¶Ç´Â ÀÌº¥Æ®"
        End Select
End Sub


'-----------------------------------------------------------------------------'
'   ±â´É : ÇØ´ç ¹®ÀÚ¿­À» ±¸ºÐÀÚ¸¦ ÀÌ¿ëÇØ ±¸ºÐÇØ ÁöÁ¤ÇÑ À§Ä¡ÀÇ ¹®ÀÚ¿­À» ±¸ÇÔ
'   ÀÎ¼ö :
'       1.pText      : ±¸ºÐÀÚ·Î ±¸¼ºµÈ ¹®ÀÚ¿­
'       2.pPosiion   : À§Ä¡
'       3.pDelimiter : ±¸ºÐÀÚ
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition ÀÎ¼ö°¡ 1ÀÎ °æ¿ì For¹® Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    'ÇØ´ç ÄÃ·³
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function


Function SetResult(asResult As String, aiItem As Integer) As String
'DB¿¡¼­ ºÒ·¯¿À±â
    Dim iFloat As Integer
    
    If Not IsNumeric(asResult) Then
        Exit Function
    End If

    Select Case aiItem
    Case 2, 16
        iFloat = 2
    Case 14, 8
        iFloat = 0
    Case Else
        iFloat = 1
    End Select

    If iFloat = 0 Then
        SetResult = CStr(CCur(asResult))
    Else
        SetResult = CStr(CCur(Left(asResult, 4 - iFloat)) & "." & Right(asResult, iFloat))
    End If
 
End Function



Private Sub optAuto_Click()
    WritePrivateProfileString "OPTION", "AutoSend", "1", App.Path & "\cp2140.ini"
End Sub

Private Sub optManual_Click()
    WritePrivateProfileString "OPTION", "AutoSend", "0", App.Path & "\cp2140.ini"
End Sub

Private Sub sspMode_Click()
    If sspMode.Caption = "¼öÁ¤¸ðµå" Then
        sspMode.Caption = "Àü¼Û¸ðµå"
        sspMode.BackColor = &HFF0000
        sspMode.ForeColor = &HFFFFFF
'        vasRes.OperationMode = 1
        
    ElseIf sspMode.Caption = "Àü¼Û¸ðµå" Then
        sspMode.Caption = "¼öÁ¤¸ðµå"
        sspMode.BackColor = &H8000&
        sspMode.ForeColor = &HFFFFFF
'        vasRes.OperationMode = 0
        
'        vasActiveCell vasRes, 1, colResult
'        vasRes.SetFocus
    End If

End Sub

Private Sub subDel_Click()
    Dim i As Long
    Dim sSpecID As String
    
    i = vasID.ActiveRow
    
    sSpecID = Trim(GetText(vasID, i, colReceNo))
    
    SQL = " Delete From pat_res " & CR & _
         " Where examdate = '" & Format(txtToday.Text, "YYYYMMDD") & "' " & CR & _
         " And equipno = '" & gEquip & "' " & CR & _
         " And diskno = '" & Trim(GetText(vasID, i, colRack)) & "' " & vbCrLf & _
         " And posno = '" & Trim(GetText(vasID, i, ColPos)) & "' " & vbCrLf & _
         " And receno = '" & Trim(GetText(vasID, i, colSampleNo)) & "' " & vbCrLf & _
         " And barcode = '" & sSpecID & "' "
    Res = SendQuery(gLocal, SQL)
    
    vasID.DeleteRows i, 1
    
End Sub

Private Sub Timer1_Timer()
        
    lngRefresh = lngRefresh - 1
    If lngRefresh = 0 Then
        Call cmdSearch_Click
        lngRefresh = gBar_Parm.WaitSec
        lblTimer.Caption = gBar_Parm.WaitSec
    Else
        lblTimer.Caption = lngRefresh
    End If

End Sub

Private Sub SpreadSheetSort(ByRef Spread As vaSpread, ByVal Col As Integer, Optional ByVal SortType As Integer = 1)
    Dim intCount As Integer
    Dim strDataField As String
    'SortType
    ' 0 : none
    ' 1 : ascending
    ' 2 : descending

    With Spread
        .Col = 1: .Col2 = .MaxCols
        .Row = 1: .Row2 = .DataRowCnt
        .SortBy = 0
        .SortKey(1) = Col       'Á¤·ÄÅ° ¿­¹øÈ£

        If SortType = 0 Then
            .SortKeyOrder(1) = SortKeyOrderNone
        ElseIf SortType = 1 Then
            .SortKeyOrder(1) = SortKeyOrderAscending
        ElseIf SortType = 2 Then
            .SortKeyOrder(1) = SortKeyOrderDescending
        Else
            .SortKeyOrder(1) = SortKeyOrderAscending
        End If

        .Action = ActionSort
    End With

End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim varTmp  As Variant
    
    If Row = 0 Then
        If Col = 1 Then
            Col = 2
        End If
        
        If OrderSort_Flag = 1 Then
            Call SpreadSheetSort(vasList, Col, 2)
            OrderSort_Flag = 2
        Else
            Call SpreadSheetSort(vasList, Col, 1)
            OrderSort_Flag = 1
        End If
    End If

End Sub


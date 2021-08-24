VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2200CD23-1176-101D-85F5-0020AF1EF604}#1.7#0"; "barcod32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmPaint 
   AutoRedraw      =   -1  'True
   Caption         =   "VB Paint"
   ClientHeight    =   10245
   ClientLeft      =   1545
   ClientTop       =   2985
   ClientWidth     =   14265
   Icon            =   "frmPaint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   14265
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   1830
      TabIndex        =   200
      Top             =   8580
      Width           =   1785
   End
   Begin VB.Frame Frame3 
      Caption         =   "Hidden Value"
      Height          =   7995
      Left            =   12090
      TabIndex        =   181
      Top             =   660
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   7890
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   194
         Text            =   "frmPaint.frx":1042
         Top             =   1200
         Width           =   3705
      End
      Begin VB.ListBox List1 
         Height          =   4920
         Left            =   60
         TabIndex        =   193
         Top             =   3000
         Width           =   6465
      End
      Begin VB.Frame Frame1 
         Height          =   2715
         Left            =   0
         TabIndex        =   182
         Top             =   270
         Width           =   6375
         Begin VB.OptionButton Option6 
            Caption         =   "Barcode"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   191
            Top             =   1830
            Width           =   1935
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Line"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   190
            Top             =   1500
            Width           =   1935
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Image"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   189
            Top             =   1170
            Width           =   1935
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Label"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   188
            Top             =   840
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmPaint.frx":11C9
            Left            =   2310
            List            =   "frmPaint.frx":11EE
            Style           =   2  '드롭다운 목록
            TabIndex        =   187
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Make"
            Height          =   405
            Left            =   2310
            TabIndex        =   186
            Top             =   810
            Width           =   1935
         End
         Begin VB.OptionButton Option2 
            Caption         =   "TextBox"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   185
            Top             =   510
            Width           =   1875
         End
         Begin VB.OptionButton Option1 
            Caption         =   "CommandButton"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   420
            TabIndex        =   184
            Top             =   210
            Width           =   1845
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  '평면
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   5220
            ScaleHeight     =   375
            ScaleWidth      =   735
            TabIndex        =   183
            Top             =   450
            Width           =   765
            Begin VB.Shape Shape1 
               BorderColor     =   &H00E0E0E0&
               Height          =   255
               Index           =   0
               Left            =   30
               Top             =   7470
               Width           =   10365
            End
         End
         Begin VB.Timer tmrMove 
            Left            =   4590
            Top             =   720
         End
         Begin BarcodLib.Barcod Barcod1 
            Height          =   315
            Left            =   2700
            TabIndex        =   192
            Tag             =   "GF07J030A195"
            Top             =   1470
            Width           =   2805
            _Version        =   65543
            _ExtentX        =   4948
            _ExtentY        =   556
            _StockProps     =   75
            Caption         =   "gf07j030a195"
            BackColor       =   16777215
            BarWidth        =   0
            Direction       =   0
            Style           =   7
            UPCNotches      =   0
            Alignment       =   0
            Extension       =   ""
         End
         Begin VB.Image Didim_DImg 
            Height          =   600
            Left            =   4260
            Top             =   2010
            Width           =   1695
         End
         Begin VB.Image Didim_SImg 
            Height          =   600
            Left            =   2490
            Top             =   2010
            Width           =   1695
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9435
      Left            =   8370
      ScaleHeight     =   627
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   697
      TabIndex        =   108
      Top             =   1110
      Width           =   10485
   End
   Begin VB.OptionButton optHW 
      Caption         =   "가로"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   25590
      TabIndex        =   107
      Top             =   1080
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.OptionButton optHW 
      Caption         =   "세로"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   26550
      TabIndex        =   106
      Top             =   1080
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtPaperHSize 
      Appearance      =   0  '평면
      Height          =   345
      Left            =   22890
      MaxLength       =   5
      TabIndex        =   105
      Text            =   "7.5"
      Top             =   690
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtPaperWSize 
      Appearance      =   0  '평면
      Height          =   345
      Left            =   24900
      MaxLength       =   5
      TabIndex        =   104
      Text            =   "3.5"
      Top             =   690
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.ComboBox cboType 
      Height          =   300
      Left            =   20010
      Style           =   2  '드롭다운 목록
      TabIndex        =   103
      Top             =   1290
      Width           =   4065
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "출력안함"
      Height          =   345
      Left            =   24150
      TabIndex        =   102
      Top             =   1290
      Width           =   1365
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  '평면
      Height          =   345
      Left            =   20010
      MaxLength       =   20
      TabIndex        =   101
      Text            =   "LINE"
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton cmdMake 
      Caption         =   "생성"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18990
      TabIndex        =   100
      Top             =   6060
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "좌표설정"
      Height          =   3255
      Left            =   25440
      TabIndex        =   93
      Top             =   2400
      Width           =   1815
      Begin VB.TextBox txtXpos 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   330
         MaxLength       =   5
         TabIndex        =   97
         Top             =   780
         Width           =   1155
      End
      Begin VB.TextBox txtYpos 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   330
         MaxLength       =   5
         TabIndex        =   96
         Top             =   1650
         Width           =   1155
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "이동"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   330
         TabIndex        =   95
         Top             =   2220
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CheckBox Check2 
         Caption         =   "미세조정"
         Height          =   345
         Left            =   270
         TabIndex        =   94
         Top             =   2790
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label9 
         Caption         =   "X 좌표"
         Height          =   285
         Left            =   360
         TabIndex        =   99
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label10 
         Caption         =   "Y 좌표"
         Height          =   285
         Left            =   360
         TabIndex        =   98
         Top             =   1320
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "적용"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   22080
      TabIndex        =   92
      Top             =   6060
      Width           =   2055
   End
   Begin VB.CommandButton cmdDelobj 
      Caption         =   "삭제"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   25110
      TabIndex        =   91
      Top             =   6060
      Width           =   2055
   End
   Begin VB.TextBox txtTag 
      Appearance      =   0  '평면
      Height          =   345
      Left            =   24090
      TabIndex        =   90
      Top             =   1680
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame Frame4 
      Caption         =   "전체좌표설정"
      Height          =   3105
      Left            =   18930
      TabIndex        =   79
      Top             =   6960
      Width           =   8235
      Begin VB.CommandButton cmdMove 
         Caption         =   "◀"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   36
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   0
         Left            =   360
         TabIndex        =   89
         Top             =   1260
         Width           =   795
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "▲"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   36
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   2
         Left            =   1170
         TabIndex        =   88
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   36
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   3
         Left            =   1170
         TabIndex        =   87
         Top             =   1980
         Width           =   795
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "▶"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   36
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   1980
         TabIndex        =   86
         Top             =   1230
         Width           =   795
      End
      Begin VB.CheckBox chkDetail 
         Caption         =   "미세조정"
         Height          =   345
         Left            =   3390
         TabIndex        =   85
         Top             =   600
         Width           =   1275
      End
      Begin VB.CheckBox chkContinue 
         Caption         =   "연속이동"
         Height          =   345
         Left            =   3390
         TabIndex        =   84
         Top             =   1050
         Width           =   1275
      End
      Begin VB.ComboBox cmbPrinter 
         Height          =   300
         Left            =   3060
         Style           =   2  '드롭다운 목록
         TabIndex        =   83
         Top             =   2370
         Width           =   4965
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "출력"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6060
         TabIndex        =   82
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox chkCorrect 
         Caption         =   "보정값적용"
         Height          =   375
         Left            =   6060
         TabIndex        =   81
         Top             =   1200
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.CheckBox chkChoice 
         Caption         =   "선택이동"
         Height          =   345
         Left            =   3390
         TabIndex        =   80
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.Frame Frame5 
      Height          =   525
      Left            =   15840
      TabIndex        =   71
      Top             =   570
      Width           =   3015
      Begin VB.OptionButton optDevide 
         Caption         =   "2배"
         Height          =   315
         Index           =   1
         Left            =   5760
         TabIndex        =   77
         Tag             =   "2"
         Top             =   180
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton optDevide 
         Caption         =   "1배"
         Height          =   315
         Index           =   0
         Left            =   3660
         TabIndex        =   76
         Tag             =   "1"
         Top             =   180
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.OptionButton optDevide 
         Caption         =   "1.4배"
         Height          =   315
         Index           =   2
         Left            =   4620
         TabIndex        =   75
         Tag             =   "1.4"
         Top             =   180
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDevide 
         Caption         =   "▶"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2130
         TabIndex        =   74
         Top             =   150
         Width           =   435
      End
      Begin VB.CommandButton cmdDevide 
         Caption         =   "◀"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   900
         TabIndex        =   73
         Top             =   150
         Width           =   435
      End
      Begin VB.TextBox txtDevide 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         TabIndex        =   72
         Top             =   150
         Width           =   765
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "배율"
         Height          =   285
         Index           =   6
         Left            =   90
         TabIndex        =   78
         Top             =   210
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "Undo"
      Height          =   405
      Left            =   19140
      TabIndex        =   70
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5520
      ScaleHeight     =   37
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   37
      TabIndex        =   65
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picImageEffect 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4650
      ScaleHeight     =   37
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   37
      TabIndex        =   61
      Top             =   300
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraTools 
      Height          =   6525
      Left            =   90
      TabIndex        =   30
      Top             =   90
      WhatsThisHelpID =   10296
      Width           =   855
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   16
         Left            =   50
         Picture         =   "frmPaint.frx":1214
         Style           =   1  '그래픽
         TabIndex        =   67
         ToolTipText     =   "Brush"
         Top             =   3120
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   17
         Left            =   435
         Picture         =   "frmPaint.frx":1378
         Style           =   1  '그래픽
         TabIndex        =   66
         ToolTipText     =   "Hand"
         Top             =   3120
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   15
         Left            =   435
         Picture         =   "frmPaint.frx":1A7A
         Style           =   1  '그래픽
         TabIndex        =   64
         ToolTipText     =   "Zoom"
         Top             =   120
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   14
         Left            =   435
         Picture         =   "frmPaint.frx":1E08
         Style           =   1  '그래픽
         TabIndex        =   63
         ToolTipText     =   "Filter Brush"
         Top             =   1245
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   13
         Left            =   435
         Picture         =   "frmPaint.frx":1E7A
         Style           =   1  '그래픽
         TabIndex        =   48
         ToolTipText     =   "Curve"
         Top             =   2745
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   12
         Left            =   435
         Picture         =   "frmPaint.frx":1ED2
         Style           =   1  '그래픽
         TabIndex        =   47
         ToolTipText     =   "Polygon"
         Top             =   2370
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   11
         Left            =   435
         Picture         =   "frmPaint.frx":1F45
         Style           =   1  '그래픽
         TabIndex        =   46
         ToolTipText     =   "Rounded Rectangle"
         Top             =   1995
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   10
         Left            =   50
         Picture         =   "frmPaint.frx":1FCF
         Style           =   1  '그래픽
         TabIndex        =   45
         ToolTipText     =   "Air Brush"
         Top             =   1245
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   2
         Left            =   435
         Picture         =   "frmPaint.frx":22D9
         Style           =   1  '그래픽
         TabIndex        =   44
         ToolTipText     =   "Eraser"
         Top             =   870
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10295
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   4
         Left            =   50
         Picture         =   "frmPaint.frx":2358
         Style           =   1  '그래픽
         TabIndex        =   43
         ToolTipText     =   "Pencil"
         Top             =   870
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         WhatsThisHelpID =   10298
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   5
         Left            =   50
         Picture         =   "frmPaint.frx":23D7
         Style           =   1  '그래픽
         TabIndex        =   42
         ToolTipText     =   "Line"
         Top             =   1620
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   3
         Left            =   435
         Picture         =   "frmPaint.frx":24BC
         Style           =   1  '그래픽
         TabIndex        =   41
         ToolTipText     =   "Fill"
         Top             =   495
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10300
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   7
         Left            =   50
         Picture         =   "frmPaint.frx":253E
         Style           =   1  '그래픽
         TabIndex        =   40
         ToolTipText     =   "Ellipse"
         Top             =   2370
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10301
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   6
         Left            =   50
         Picture         =   "frmPaint.frx":25AB
         Style           =   1  '그래픽
         TabIndex        =   39
         ToolTipText     =   "Rectangle"
         Top             =   1995
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10302
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   8
         Left            =   50
         Picture         =   "frmPaint.frx":2618
         Style           =   1  '그래픽
         TabIndex        =   38
         ToolTipText     =   "Text"
         Top             =   2745
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   9
         Left            =   435
         Picture         =   "frmPaint.frx":299A
         Style           =   1  '그래픽
         TabIndex        =   37
         ToolTipText     =   "Arrow"
         Top             =   1620
         WhatsThisHelpID =   10340
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   0
         Left            =   50
         Picture         =   "frmPaint.frx":29E5
         Style           =   1  '그래픽
         TabIndex        =   36
         ToolTipText     =   "Select Area"
         Top             =   120
         WhatsThisHelpID =   10359
         Width           =   390
      End
      Begin VB.OptionButton optTools 
         Height          =   375
         Index           =   1
         Left            =   50
         Picture         =   "frmPaint.frx":2D63
         Style           =   1  '그래픽
         TabIndex        =   35
         ToolTipText     =   "Pick Color"
         Top             =   495
         WhatsThisHelpID =   10361
         Width           =   390
      End
      Begin VB.Frame fraOptDot 
         Height          =   1215
         Left            =   90
         TabIndex        =   31
         Top             =   3600
         WhatsThisHelpID =   10335
         Width           =   660
         Begin VB.Label lblDot 
            BackColor       =   &H8000000D&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            Height          =   255
            Left            =   75
            TabIndex        =   32
            Top             =   150
            WhatsThisHelpID =   10336
            Width           =   255
         End
         Begin VB.Shape shpDot 
            BorderStyle     =   0  '투명
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  '단색
            Height          =   30
            Index           =   0
            Left            =   195
            Shape           =   3  '원형
            Top             =   270
            Width           =   30
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  '단색
            Height          =   45
            Index           =   1
            Left            =   435
            Shape           =   3  '원형
            Top             =   255
            Width           =   45
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  '단색
            Height          =   60
            Index           =   2
            Left            =   165
            Shape           =   3  '원형
            Top             =   495
            Width           =   60
         End
         Begin VB.Shape shpDot 
            BorderStyle     =   0  '투명
            FillStyle       =   0  '단색
            Height          =   75
            Index           =   3
            Left            =   420
            Shape           =   3  '원형
            Top             =   495
            Width           =   75
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  '단색
            Height          =   90
            Index           =   4
            Left            =   150
            Shape           =   3  '원형
            Top             =   730
            Width           =   90
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  '단색
            Height          =   105
            Index           =   5
            Left            =   405
            Shape           =   3  '원형
            Top             =   715
            Width           =   105
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  '단색
            Height          =   120
            Index           =   6
            Left            =   140
            Shape           =   3  '원형
            Top             =   970
            Width           =   120
         End
         Begin VB.Shape shpDot 
            FillStyle       =   0  '단색
            Height          =   135
            Index           =   7
            Left            =   390
            Shape           =   3  '원형
            Top             =   960
            Width           =   135
         End
      End
      Begin VB.Frame fraBrush 
         Height          =   1545
         Left            =   90
         TabIndex        =   68
         Top             =   4815
         Visible         =   0   'False
         WhatsThisHelpID =   10335
         Width           =   660
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   9
            Left            =   405
            Picture         =   "frmPaint.frx":30EB
            Top             =   1290
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   8
            Left            =   120
            Picture         =   "frmPaint.frx":312D
            Top             =   1290
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   1
            Left            =   405
            Picture         =   "frmPaint.frx":316C
            Top             =   210
            Width           =   135
         End
         Begin VB.Label lblBrush 
            BackColor       =   &H8000000D&
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            Height          =   255
            Left            =   60
            TabIndex        =   69
            Top             =   150
            WhatsThisHelpID =   10336
            Width           =   255
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   0
            Left            =   120
            Picture         =   "frmPaint.frx":31B0
            Top             =   210
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   3
            Left            =   405
            Picture         =   "frmPaint.frx":31F4
            Top             =   480
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   2
            Left            =   120
            Picture         =   "frmPaint.frx":3238
            Top             =   480
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   6
            Left            =   120
            Picture         =   "frmPaint.frx":327D
            Top             =   1020
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   7
            Left            =   405
            Picture         =   "frmPaint.frx":32BF
            Top             =   1020
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   5
            Left            =   405
            Picture         =   "frmPaint.frx":3301
            Top             =   750
            Width           =   135
         End
         Begin VB.Image imgBrush 
            Appearance      =   0  '평면
            Height          =   135
            Index           =   4
            Left            =   120
            Picture         =   "frmPaint.frx":3345
            Top             =   750
            Width           =   135
         End
      End
      Begin VB.Frame fraOptFill 
         Height          =   1110
         Left            =   75
         TabIndex        =   33
         Top             =   4815
         Visible         =   0   'False
         WhatsThisHelpID =   10333
         Width           =   705
         Begin VB.Label lblFill 
            BackStyle       =   0  '투명
            BorderStyle     =   1  '단일 고정
            Height          =   275
            Left            =   60
            TabIndex        =   34
            Top             =   150
            WhatsThisHelpID =   10334
            Width           =   570
         End
         Begin VB.Shape shpRect 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FFFFFF&
            Height          =   150
            Index           =   0
            Left            =   140
            Top             =   210
            Width           =   420
         End
         Begin VB.Shape shpRect 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   150
            Index           =   1
            Left            =   135
            Top             =   525
            Width           =   420
         End
         Begin VB.Shape shpRect 
            BorderStyle     =   0  '투명
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   150
            Index           =   2
            Left            =   140
            Top             =   840
            Width           =   420
         End
      End
   End
   Begin VB.HScrollBar hscPaint 
      Height          =   255
      LargeChange     =   100
      Left            =   885
      Max             =   0
      SmallChange     =   10
      TabIndex        =   55
      Top             =   6330
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.VScrollBar vscPaint 
      Height          =   6165
      LargeChange     =   1000
      Left            =   7650
      Max             =   0
      SmallChange     =   100
      TabIndex        =   56
      Top             =   270
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraScroll 
      BorderStyle     =   0  '없음
      Height          =   255
      Left            =   7230
      TabIndex        =   57
      Top             =   5850
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraColor 
      Height          =   860
      Left            =   30
      TabIndex        =   0
      Top             =   6870
      Width           =   7455
      Begin MSComDlg.CommonDialog cdlPrint 
         Left            =   4680
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin MSComDlg.CommonDialog cdlFonts 
         Left            =   5190
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdlOpen 
         Left            =   6210
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   $"frmPaint.frx":3388
         Flags           =   4
      End
      Begin MSComDlg.CommonDialog cdlColor 
         Left            =   5715
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdlSave 
         Left            =   6720
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "*.brg"
         DialogTitle     =   "Save As"
         Filter          =   "Bitmap Files (*.bmp) |*.bmp"
      End
      Begin VB.Label lblColor 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   25
         Left            =   4080
         TabIndex        =   29
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00004080&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   24
         Left            =   4080
         TabIndex        =   28
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   23
         Left            =   3825
         TabIndex        =   27
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00400040&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   22
         Left            =   3825
         TabIndex        =   26
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   21
         Left            =   3555
         TabIndex        =   25
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   20
         Left            =   3555
         TabIndex        =   24
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   19
         Left            =   3285
         TabIndex        =   23
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00004000&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   18
         Left            =   3285
         TabIndex        =   22
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   17
         Left            =   3015
         TabIndex        =   21
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00004040&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   16
         Left            =   3015
         TabIndex        =   20
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   15
         Left            =   2745
         TabIndex        =   19
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00800080&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   14
         Left            =   2745
         TabIndex        =   18
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   13
         Left            =   2475
         TabIndex        =   17
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00800000&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   12
         Left            =   2475
         TabIndex        =   16
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   11
         Left            =   2200
         TabIndex        =   15
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00808000&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   10
         Left            =   2200
         TabIndex        =   14
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   9
         Left            =   1935
         TabIndex        =   13
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00008000&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   8
         Left            =   1935
         TabIndex        =   12
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   7
         Left            =   1660
         TabIndex        =   11
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00008080&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   6
         Left            =   1660
         TabIndex        =   10
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   5
         Left            =   1400
         TabIndex        =   9
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000080&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   4
         Left            =   1400
         TabIndex        =   8
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   3
         Left            =   1125
         TabIndex        =   7
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00808080&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   2
         Left            =   1125
         TabIndex        =   6
         Top             =   225
         Width           =   255
      End
      Begin VB.Label lblForeColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Left            =   255
         TabIndex        =   4
         Top             =   300
         Width           =   255
      End
      Begin VB.Label lblFillColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Left            =   375
         TabIndex        =   5
         Top             =   420
         Width           =   255
      End
      Begin VB.Label label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  '단일 고정
         Height          =   555
         Left            =   150
         TabIndex        =   3
         Top             =   210
         Width           =   555
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   1
         Left            =   850
         TabIndex        =   2
         Top             =   495
         Width           =   255
      End
      Begin VB.Label lblColor 
         BackColor       =   &H00000000&
         BorderStyle     =   1  '단일 고정
         Height          =   255
         Index           =   0
         Left            =   850
         TabIndex        =   1
         Top             =   225
         Width           =   255
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  '아래 맞춤
      Height          =   255
      Left            =   0
      TabIndex        =   54
      Top             =   9990
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21063
            MinWidth        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPaint 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6030
      Left            =   930
      MousePointer    =   99  '사용자 정의
      ScaleHeight     =   398
      ScaleMode       =   3  '픽셀
      ScaleWidth      =   411
      TabIndex        =   49
      Top             =   195
      Width           =   6225
      Begin VB.PictureBox picClipboard 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Height          =   630
         Left            =   1200
         ScaleHeight     =   42
         ScaleMode       =   3  '픽셀
         ScaleWidth      =   41
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtText 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   52
         Top             =   180
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.PictureBox picBuffer 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   2040
         ScaleHeight     =   37
         ScaleMode       =   3  '픽셀
         ScaleWidth      =   37
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox picSelect 
         Appearance      =   0  '평면
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   480
         ScaleHeight     =   42
         ScaleMode       =   3  '픽셀
         ScaleWidth      =   41
         TabIndex        =   50
         Top             =   120
         Width           =   615
      End
      Begin VB.Image imgBezier 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   60
         Index           =   0
         Left            =   2880
         Top             =   240
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgBezier 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   60
         Index           =   3
         Left            =   3240
         Top             =   600
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgBezier 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   60
         Index           =   2
         Left            =   3240
         Top             =   240
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgBezier 
         Appearance      =   0  '평면
         BorderStyle     =   1  '단일 고정
         Height          =   60
         Index           =   1
         Left            =   2880
         Top             =   600
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label lblTextSize 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   330
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.PictureBox picPaintResize 
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Index           =   0
      Left            =   7110
      MousePointer    =   9  'W E 크기 조정
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   58
      Top             =   3180
      Width           =   70
   End
   Begin VB.PictureBox picPaintResize 
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Index           =   2
      Left            =   7110
      MousePointer    =   8  'NW SE 크기 조정
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   60
      Top             =   6225
      Width           =   70
   End
   Begin VB.PictureBox picPaintResize 
      Appearance      =   0  '평면
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Index           =   1
      Left            =   3960
      MousePointer    =   7  'N S크기 조정
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   59
      Top             =   6225
      Width           =   70
   End
   Begin TabDlg.SSTab sstType 
      Height          =   3315
      Left            =   18930
      TabIndex        =   109
      Top             =   2400
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   5847
      _Version        =   393216
      Tabs            =   6
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "S_Text"
      TabPicture(0)   =   "frmPaint.frx":347C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label8(0)"
      Tab(0).Control(1)=   "Label7(0)"
      Tab(0).Control(2)=   "Label6(0)"
      Tab(0).Control(3)=   "Label8(6)"
      Tab(0).Control(4)=   "cmdFont(0)"
      Tab(0).Control(5)=   "chkTStatic"
      Tab(0).Control(6)=   "txtContent(0)"
      Tab(0).Control(7)=   "chkFontItalic(0)"
      Tab(0).Control(8)=   "chkFontUnder(0)"
      Tab(0).Control(9)=   "chkFontBold(0)"
      Tab(0).Control(10)=   "txtFontName(0)"
      Tab(0).Control(11)=   "txtFontSize(0)"
      Tab(0).Control(12)=   "optSTRotate(0)"
      Tab(0).Control(13)=   "optSTRotate(1)"
      Tab(0).Control(14)=   "optSTRotate(2)"
      Tab(0).Control(15)=   "optSTRotate(3)"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "D_Text"
      TabPicture(1)   =   "frmPaint.frx":3498
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6(1)"
      Tab(1).Control(1)=   "Label7(1)"
      Tab(1).Control(2)=   "Label8(1)"
      Tab(1).Control(3)=   "txtFontSize(1)"
      Tab(1).Control(4)=   "txtFontName(1)"
      Tab(1).Control(5)=   "chkFontBold(1)"
      Tab(1).Control(6)=   "chkFontUnder(1)"
      Tab(1).Control(7)=   "chkFontItalic(1)"
      Tab(1).Control(8)=   "txtContent(1)"
      Tab(1).Control(9)=   "cmdFont(1)"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "S_Image"
      TabPicture(2)   =   "frmPaint.frx":34B4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtImageWSize(2)"
      Tab(2).Control(1)=   "txtImageHSize(2)"
      Tab(2).Control(2)=   "cmdImageDevSet(0)"
      Tab(2).Control(3)=   "txtImageDevide(0)"
      Tab(2).Control(4)=   "chkIStatic"
      Tab(2).Control(5)=   "txtImageHSize(0)"
      Tab(2).Control(6)=   "txtImageName(0)"
      Tab(2).Control(7)=   "txtImageWSize(0)"
      Tab(2).Control(8)=   "cmdImage(0)"
      Tab(2).Control(9)=   "Label8(8)"
      Tab(2).Control(10)=   "Label8(7)"
      Tab(2).Control(11)=   "Label8(2)"
      Tab(2).Control(12)=   "Label7(2)"
      Tab(2).Control(13)=   "Label6(2)"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "D_Image"
      TabPicture(3)   =   "frmPaint.frx":34D0
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label8(10)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label8(9)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label8(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label7(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label6(3)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtImageWSize(3)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtImageHSize(3)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdImageDevSet(1)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtImageDevide(1)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txtImageHSize(1)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "txtImageName(1)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "txtImageWSize(1)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "cmdImage(1)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).ControlCount=   13
      TabCaption(4)   =   "Barcode"
      TabPicture(4)   =   "frmPaint.frx":34EC
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chkBarRotate"
      Tab(4).Control(1)=   "txtBarData"
      Tab(4).Control(2)=   "txtBarDevide"
      Tab(4).Control(3)=   "txtBarWSize"
      Tab(4).Control(4)=   "txtBarHSize"
      Tab(4).Control(5)=   "cboBarType"
      Tab(4).Control(6)=   "Label8(4)"
      Tab(4).Control(7)=   "Label7(4)"
      Tab(4).Control(8)=   "Label6(4)"
      Tab(4).Control(9)=   "Label8(5)"
      Tab(4).Control(10)=   "Label7(5)"
      Tab(4).ControlCount=   11
      TabCaption(5)   =   "Line"
      TabPicture(5)   =   "frmPaint.frx":3508
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtLineWSize"
      Tab(5).Control(1)=   "txtLineHSize"
      Tab(5).Control(2)=   "chkLineRotate"
      Tab(5).Control(3)=   "Label6(5)"
      Tab(5).Control(4)=   "Label7(6)"
      Tab(5).ControlCount=   5
      Begin VB.CheckBox chkLineRotate 
         Caption         =   "Portrait"
         Height          =   345
         Left            =   -72480
         TabIndex        =   154
         Top             =   2010
         Width           =   1275
      End
      Begin VB.TextBox txtLineHSize 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   -72480
         MaxLength       =   1
         TabIndex        =   153
         Top             =   990
         Width           =   3225
      End
      Begin VB.TextBox txtLineWSize 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   -72480
         MaxLength       =   5
         TabIndex        =   152
         Top             =   1500
         Width           =   3225
      End
      Begin VB.ComboBox cboBarType 
         Height          =   300
         Left            =   -72480
         Style           =   2  '드롭다운 목록
         TabIndex        =   151
         Top             =   600
         Width           =   3225
      End
      Begin VB.TextBox txtBarHSize 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   -72480
         MaxLength       =   5
         TabIndex        =   150
         Top             =   1920
         Width           =   3225
      End
      Begin VB.TextBox txtBarWSize 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   -72480
         MaxLength       =   5
         TabIndex        =   149
         Top             =   1500
         Width           =   3225
      End
      Begin VB.TextBox txtBarDevide 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   -72480
         MaxLength       =   1
         TabIndex        =   148
         Top             =   1080
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.TextBox txtBarData 
         Appearance      =   0  '평면
         Height          =   345
         Left            =   -72480
         MaxLength       =   20
         TabIndex        =   147
         Top             =   2340
         Width           =   3225
      End
      Begin VB.CheckBox chkBarRotate 
         Caption         =   "Portrait"
         Height          =   345
         Left            =   -72450
         TabIndex        =   146
         Top             =   2820
         Width           =   1665
      End
      Begin VB.CommandButton cmdImage 
         Caption         =   "이미지 설정"
         Height          =   405
         Index           =   1
         Left            =   3810
         TabIndex        =   145
         Top             =   2730
         Width           =   1935
      End
      Begin VB.TextBox txtImageWSize 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   1
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   144
         Top             =   1410
         Width           =   1605
      End
      Begin VB.TextBox txtImageName 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   1
         Left            =   480
         TabIndex        =   143
         Top             =   930
         Width           =   5265
      End
      Begin VB.TextBox txtImageHSize 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   1
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   142
         Top             =   1830
         Width           =   1605
      End
      Begin VB.CommandButton cmdImage 
         Caption         =   "이미지 설정"
         Height          =   405
         Index           =   0
         Left            =   -71190
         TabIndex        =   141
         Top             =   2730
         Width           =   1935
      End
      Begin VB.TextBox txtImageWSize 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   0
         Left            =   -72480
         MaxLength       =   5
         TabIndex        =   140
         Top             =   1410
         Width           =   1605
      End
      Begin VB.TextBox txtImageName 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   0
         Left            =   -74520
         TabIndex        =   139
         Top             =   930
         Width           =   5265
      End
      Begin VB.TextBox txtImageHSize 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   0
         Left            =   -72480
         MaxLength       =   5
         TabIndex        =   138
         Top             =   1830
         Width           =   1605
      End
      Begin VB.CheckBox chkIStatic 
         Caption         =   "무조건 고정"
         Height          =   345
         Left            =   -74190
         TabIndex        =   137
         Top             =   2820
         Width           =   1665
      End
      Begin VB.TextBox txtFontSize 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   1
         Left            =   -72480
         MaxLength       =   3
         TabIndex        =   136
         Top             =   960
         Width           =   3225
      End
      Begin VB.TextBox txtFontName 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   1
         Left            =   -72480
         MaxLength       =   20
         TabIndex        =   135
         Top             =   540
         Width           =   3225
      End
      Begin VB.CheckBox chkFontBold 
         Caption         =   "굵게"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   -72480
         TabIndex        =   134
         Top             =   1380
         Width           =   825
      End
      Begin VB.CheckBox chkFontUnder 
         Caption         =   "밑줄"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   -71490
         TabIndex        =   133
         Top             =   1380
         Width           =   825
      End
      Begin VB.CheckBox chkFontItalic 
         Caption         =   "기울게 "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   -70530
         TabIndex        =   132
         Top             =   1380
         Width           =   1065
      End
      Begin VB.TextBox txtContent 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   1
         Left            =   -72480
         TabIndex        =   131
         Top             =   1770
         Width           =   3225
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "Font 설정"
         Height          =   405
         Index           =   1
         Left            =   -71190
         TabIndex        =   130
         Top             =   2730
         Width           =   1935
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "Font 설정"
         Height          =   405
         Index           =   0
         Left            =   -71190
         TabIndex        =   129
         Top             =   2730
         Width           =   1935
      End
      Begin VB.CheckBox chkTStatic 
         Caption         =   "무조건 고정"
         Height          =   345
         Left            =   -74190
         TabIndex        =   128
         Top             =   2820
         Width           =   1665
      End
      Begin VB.TextBox txtContent 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   0
         Left            =   -72480
         TabIndex        =   127
         Top             =   1770
         Width           =   3225
      End
      Begin VB.CheckBox chkFontItalic 
         Caption         =   "기울게 "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   -70530
         TabIndex        =   126
         Top             =   1380
         Width           =   1155
      End
      Begin VB.CheckBox chkFontUnder 
         Caption         =   "밑줄"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   -71490
         TabIndex        =   125
         Top             =   1380
         Width           =   825
      End
      Begin VB.CheckBox chkFontBold 
         Caption         =   "굵게"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   -72480
         TabIndex        =   124
         Top             =   1380
         Width           =   825
      End
      Begin VB.TextBox txtFontName 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   0
         Left            =   -72480
         MaxLength       =   20
         TabIndex        =   123
         Top             =   540
         Width           =   3225
      End
      Begin VB.TextBox txtFontSize 
         Appearance      =   0  '평면
         Height          =   345
         Index           =   0
         Left            =   -72480
         MaxLength       =   3
         TabIndex        =   122
         Top             =   960
         Width           =   3225
      End
      Begin VB.OptionButton optSTRotate 
         Caption         =   "0˚"
         Height          =   255
         Index           =   0
         Left            =   -72450
         TabIndex        =   121
         Top             =   2250
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.OptionButton optSTRotate 
         Caption         =   "90˚"
         Height          =   255
         Index           =   1
         Left            =   -71670
         TabIndex        =   120
         Top             =   2250
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.OptionButton optSTRotate 
         Caption         =   "180˚"
         Height          =   255
         Index           =   2
         Left            =   -70860
         TabIndex        =   119
         Top             =   2250
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.OptionButton optSTRotate 
         Caption         =   "270˚"
         Height          =   255
         Index           =   3
         Left            =   -69960
         TabIndex        =   118
         Top             =   2250
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtImageDevide 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   345
         Index           =   0
         Left            =   -72480
         MaxLength       =   5
         TabIndex        =   117
         Top             =   2250
         Width           =   585
      End
      Begin VB.CommandButton cmdImageDevSet 
         Caption         =   "적용"
         Height          =   375
         Index           =   0
         Left            =   -71550
         TabIndex        =   116
         Top             =   2250
         Width           =   585
      End
      Begin VB.TextBox txtImageDevide 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Height          =   345
         Index           =   1
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   115
         Top             =   2250
         Width           =   585
      End
      Begin VB.CommandButton cmdImageDevSet 
         Caption         =   "적용"
         Height          =   375
         Index           =   1
         Left            =   3450
         TabIndex        =   114
         Top             =   2250
         Width           =   585
      End
      Begin VB.TextBox txtImageHSize 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   345
         Index           =   2
         Left            =   -70860
         MaxLength       =   5
         TabIndex        =   113
         Top             =   1830
         Width           =   1605
      End
      Begin VB.TextBox txtImageWSize 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   345
         Index           =   2
         Left            =   -70860
         MaxLength       =   5
         TabIndex        =   112
         Top             =   1410
         Width           =   1605
      End
      Begin VB.TextBox txtImageHSize 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   345
         Index           =   3
         Left            =   4140
         MaxLength       =   5
         TabIndex        =   111
         Top             =   1830
         Width           =   1605
      End
      Begin VB.TextBox txtImageWSize 
         Appearance      =   0  '평면
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   345
         Index           =   3
         Left            =   4140
         MaxLength       =   5
         TabIndex        =   110
         Top             =   1410
         Width           =   1605
      End
      Begin VB.Label Label7 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "선길이 "
         Height          =   285
         Index           =   6
         Left            =   -74730
         TabIndex        =   178
         Top             =   1530
         Width           =   1635
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "선굵기 "
         Height          =   285
         Index           =   5
         Left            =   -74730
         TabIndex        =   177
         Top             =   1050
         Width           =   1635
      End
      Begin VB.Label Label7 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "Data "
         Height          =   285
         Index           =   5
         Left            =   -74730
         TabIndex        =   176
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "길이Size "
         Height          =   285
         Index           =   5
         Left            =   -74730
         TabIndex        =   175
         Top             =   1530
         Width           =   1635
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "종류 "
         Height          =   285
         Index           =   4
         Left            =   -74730
         TabIndex        =   174
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label Label7 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "간격비율 "
         Height          =   285
         Index           =   4
         Left            =   -74730
         TabIndex        =   173
         Top             =   1110
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "높이Size "
         Height          =   285
         Index           =   4
         Left            =   -74730
         TabIndex        =   172
         Top             =   1980
         Width           =   1635
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "이미지"
         Height          =   285
         Index           =   3
         Left            =   270
         TabIndex        =   171
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "가로Size "
         Height          =   285
         Index           =   3
         Left            =   1020
         TabIndex        =   170
         Top             =   1470
         Width           =   1185
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "세로Size"
         Height          =   285
         Index           =   3
         Left            =   1020
         TabIndex        =   169
         Top             =   1890
         Width           =   1125
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "이미지"
         Height          =   285
         Index           =   2
         Left            =   -74730
         TabIndex        =   168
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "가로Size "
         Height          =   285
         Index           =   2
         Left            =   -73980
         TabIndex        =   167
         Top             =   1470
         Width           =   1185
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "세로Size"
         Height          =   285
         Index           =   2
         Left            =   -73980
         TabIndex        =   166
         Top             =   1890
         Width           =   1125
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "Font Name "
         Height          =   285
         Index           =   1
         Left            =   -74730
         TabIndex        =   165
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label7 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "Font Size "
         Height          =   285
         Index           =   1
         Left            =   -74730
         TabIndex        =   164
         Top             =   990
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "Text "
         Height          =   285
         Index           =   1
         Left            =   -74730
         TabIndex        =   163
         Top             =   1830
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "Text "
         Height          =   285
         Index           =   0
         Left            =   -74730
         TabIndex        =   162
         Top             =   1830
         Width           =   1635
      End
      Begin VB.Label Label7 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "Font Size "
         Height          =   285
         Index           =   0
         Left            =   -74730
         TabIndex        =   161
         Top             =   990
         Width           =   1635
      End
      Begin VB.Label Label6 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "Font Name "
         Height          =   285
         Index           =   0
         Left            =   -74730
         TabIndex        =   160
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "회전"
         Height          =   285
         Index           =   6
         Left            =   -74730
         TabIndex        =   159
         Top             =   2250
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "배율"
         Height          =   285
         Index           =   7
         Left            =   -73980
         TabIndex        =   158
         Top             =   2310
         Width           =   1125
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "%"
         Height          =   285
         Index           =   8
         Left            =   -71970
         TabIndex        =   157
         Top             =   2310
         Width           =   315
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "배율"
         Height          =   285
         Index           =   9
         Left            =   1020
         TabIndex        =   156
         Top             =   2310
         Width           =   1125
      End
      Begin VB.Label Label8 
         Alignment       =   1  '오른쪽 맞춤
         Caption         =   "%"
         Height          =   285
         Index           =   10
         Left            =   3030
         TabIndex        =   155
         Top             =   2310
         Width           =   315
      End
   End
   Begin FPSpread.vaSpread spdList 
      Height          =   3585
      Left            =   8400
      TabIndex        =   179
      Top             =   10560
      Width           =   18915
      _Version        =   196608
      _ExtentX        =   33364
      _ExtentY        =   6324
      _StockProps     =   64
      AutoCalc        =   0   'False
      AutoClipboard   =   0   'False
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      ColsFrozen      =   3
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FormulaSync     =   0   'False
      GridShowHoriz   =   0   'False
      GridSolid       =   0   'False
      MaxCols         =   29
      MaxRows         =   5
      MoveActiveOnFocus=   0   'False
      OperationMode   =   2
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      SelectBlockOptions=   0
      ShadowColor     =   14735309
      SpreadDesigner  =   "frmPaint.frx":3524
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  '위 맞춤
      Height          =   555
      Left            =   0
      TabIndex        =   180
      Top             =   0
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   979
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlToolbar 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPaint.frx":3F9B
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPaint.frx":4C8D
               Key             =   "Make"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPaint.frx":A8AF
               Key             =   "View"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPaint.frx":C589
               Key             =   "Exit"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPaint.frx":122EB
               Key             =   "Edit"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPaint.frx":13FC5
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPaint.frx":15C9F
               Key             =   "New"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   8430
      Picture         =   "frmPaint.frx":17979
      Top             =   660
      Width           =   3540
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "용지설정(높이X넓이)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   20940
      TabIndex        =   199
      Top             =   780
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "cm  X"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   24210
      TabIndex        =   198
      Top             =   750
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "cm"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   26220
      TabIndex        =   197
      Top             =   750
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "항목명 "
      Height          =   285
      Left            =   18990
      TabIndex        =   196
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "구분 "
      Height          =   285
      Left            =   18990
      TabIndex        =   195
      Top             =   1320
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "C&rop"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Begin VB.Menu mnuBorderStyle 
         Caption         =   "&Border Style"
         Begin VB.Menu mnuBS 
            Caption         =   "&Solid"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuBS 
            Caption         =   "&Dash"
            Index           =   1
         End
         Begin VB.Menu mnuBS 
            Caption         =   "D&ot"
            Index           =   2
         End
         Begin VB.Menu mnuBS 
            Caption         =   "D&ashDot"
            Index           =   3
         End
         Begin VB.Menu mnuBS 
            Caption         =   "Da&shDotDot"
            Index           =   4
         End
      End
      Begin VB.Menu mnuFillStyle 
         Caption         =   "Fi&ll Style"
         Begin VB.Menu mnuFS 
            Caption         =   "&Solid"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Transparent"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Horizontal Line"
            Index           =   2
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Vertical Line"
            Index           =   3
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Downward Diagonal"
            Index           =   4
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Upward Diagonal"
            Index           =   5
         End
         Begin VB.Menu mnuFS 
            Caption         =   "&Cross"
            Index           =   6
         End
         Begin VB.Menu mnuFS 
            Caption         =   "Diagona&l Cross"
            Index           =   7
         End
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuForegroundColor 
         Caption         =   "F&oreground Color..."
      End
      Begin VB.Menu mnuFillColor 
         Caption         =   "Fi&ll Color..."
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font..."
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuEffect 
      Caption         =   "Effec&t"
      Begin VB.Menu mnuResize 
         Caption         =   "Re&size"
         Begin VB.Menu mnuResize25 
            Caption         =   "25%"
         End
         Begin VB.Menu mnuResize50 
            Caption         =   "50%"
         End
         Begin VB.Menu mnuResize75 
            Caption         =   "75%"
         End
         Begin VB.Menu mnuResize125 
            Caption         =   "125%"
         End
         Begin VB.Menu mnuResize150 
            Caption         =   "150%"
         End
         Begin VB.Menu mnuResize175 
            Caption         =   "175%"
         End
         Begin VB.Menu mnuResize200 
            Caption         =   "200%"
         End
         Begin VB.Menu mnuSep6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuResizeBoth 
            Caption         =   "&Both"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuResizeWidth 
            Caption         =   "&Width"
         End
         Begin VB.Menu mnuResizeHeight 
            Caption         =   "&Height"
         End
      End
      Begin VB.Menu mnuFlip 
         Caption         =   "&Flip"
         Begin VB.Menu mnuFlipHorizontal 
            Caption         =   "&Horizontal"
         End
         Begin VB.Menu mnuFlipVertical 
            Caption         =   "&Vertical"
         End
      End
      Begin VB.Menu mnuRotate 
         Caption         =   "&Rotate"
         Begin VB.Menu mnuRotate45 
            Caption         =   $"frmPaint.frx":1C72B
         End
         Begin VB.Menu mnuRotate90 
            Caption         =   $"frmPaint.frx":1C735
         End
         Begin VB.Menu mnuRotate135 
            Caption         =   $"frmPaint.frx":1C73F
         End
         Begin VB.Menu mnuRotate180 
            Caption         =   $"frmPaint.frx":1C74A
         End
         Begin VB.Menu mnuRotate225 
            Caption         =   $"frmPaint.frx":1C755
         End
         Begin VB.Menu mnuRotate270 
            Caption         =   $"frmPaint.frx":1C760
         End
         Begin VB.Menu mnuRotate315 
            Caption         =   $"frmPaint.frx":1C76B
         End
         Begin VB.Menu mnuSep7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRotateClockwise 
            Caption         =   "&Clockwise"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRotateAntiClockwise 
            Caption         =   "&Anti-Clockwise"
         End
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "&Filte&r"
      Begin VB.Menu mnuBlacknWhite 
         Caption         =   "&Black && White"
      End
      Begin VB.Menu mnuBlur 
         Caption         =   "B&lur"
      End
      Begin VB.Menu mnuBrightness 
         Caption         =   "B&rightness"
      End
      Begin VB.Menu mnuCrease 
         Caption         =   "&Crease"
      End
      Begin VB.Menu mnuDarkness 
         Caption         =   "&Darkness"
      End
      Begin VB.Menu mnuDiffuse 
         Caption         =   "Di&ffuse"
      End
      Begin VB.Menu mnuEmboss 
         Caption         =   "&Emboss"
      End
      Begin VB.Menu mnuGrayBlacknWhite 
         Caption         =   "Gra&y Black && White"
      End
      Begin VB.Menu mnuGrayscale 
         Caption         =   "&Grayscale"
      End
      Begin VB.Menu mnuInvertColors 
         Caption         =   "&Invert Colors"
      End
      Begin VB.Menu mnuReplaceColors 
         Caption         =   "&Replace Colors"
      End
      Begin VB.Menu mnuSharpen 
         Caption         =   "&Sharpen"
      End
      Begin VB.Menu mnuSnow 
         Caption         =   "S&now"
      End
      Begin VB.Menu mnuWave 
         Caption         =   "&Wave"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuTFilter 
      Caption         =   "&TFilter"
      Visible         =   0   'False
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Black && White"
         Index           =   0
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "B&lur"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Light"
         Index           =   2
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Crease"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Dirty"
         Index           =   4
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "Di&ffuse"
         Index           =   5
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Emboss"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "Gra&y Black && White"
         Index           =   7
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Grayscale"
         Index           =   8
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Invert Colors"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Replace Color"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Sharpen"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "S&now"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilterTools 
         Caption         =   "&Wave"
         Index           =   13
      End
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name  : frmPaint.frm                                                 **
'** Language   : Visual Basic 6.0                                             **
'** References : Microsoft Scripting Runtime (only for mdlGeneral.ForceSave)  **
'** Components : * Microsoft Common Dialog Control 6.0 (SP3)                  **
'**              * Microsoft Windows Common Controls 6.0                      **
'** Modules    : mdlAPI, mdlEffect, mdlFilter, mdlGeneral                     **
'** Developer  : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A simple drawing program similar to Microsoft Paint plus     **
'**              several image filters                                        **
'** Features   :                                                              **
'** - Drawing tools: curve, polygon, filter brush, brush (10 different        **
'**                  shapes), air brush, text, fill, rectangle, square,       **
'**                  rounded rectangle, rounded square, ellipse, circle,      **
'**                  pencil, eraser and pick                                  **
'** - Drawing properties: foreground color, fill color, fill style,           **
'**                       draw width, border style and font                   **
'** - Selection tool: move, cut, copy, paste, delete, crop, apply effects,    **
'**                   apply filters                                           **
'** - Effects: resize, flip horizontal/vertical, rotate, clear                **
'** - Filters: black and white, blur, brightness, crease, darkness, diffuse,  **
'**            emboss, gray black and white, grayscale, invert colors,        **
'**            replace colors, sharpen, snow and wave                         **
'** - Undo/redo (limited only by memory, currently I set it to 10x undo/redo) **
'** - Others: scroll bars, zoom, resizable paint area, hand, status bar,      **
'**           open, save, and print                                           **
'** Version    : 1.02                                                         **
'** - 1.00 (August 10, 2003)                                                  **
'** - 1.01 (August 13, 2003):                                                 **
'**     * bugs fixed on pressing cancel on open, save as and print dialog box **
'**     * bugs fixed on filter brushing on the top left of the paint area     **
'** - 1.02 (August 15, 2003):                                                 **
'**     bugs fixed on resizing and zooming the image multiple times           **
'** - 1.03 (September 12, 2003):                                              **
'**     turns on several error-handler                                        **
'** Last modified on September 12, 2003                                       **
'*******************************************************************************

Option Explicit

'Enumeration declaration
Public Enum enmStatusBar
  conStPaintArea = 0
  conStColorBox = 1
  conStForeColorBox = 2
  conStBackColorBox = 3
  conStFiltering = 4
  conStRetrieveingColor = 5
End Enum
Dim sng As Single

Private Enum enmTool
  'the values below must match optTools index
  conTSelect = 0
  conTPick = 1
  conTEraser = 2
  conTFill = 3
  conTPencil = 4
  conTLine = 5
  conTRect = 6
  conTEllipse = 7
  conTText = 8
  conTArrow = 9
  conTAirBrush = 10
  conTRoundRect = 11
  conTPolygon = 12
  conTCurve = 13
  conTFilter = 14
  conTZoom = 15
  conTBrush = 16
  conTHand = 17
End Enum

Private Enum enmFillStyle
  conTsBorderOnly = 0
  conTsBorderFill = 1
  conTsFillOnly = 2
End Enum

Private Enum enmBrushShape
  'the values below must match imgBrush index
  conFilledRect = 0
  conFilledCircle = 1
  conRect = 2
  conCircle = 3
  conCross = 4
  conDiagonalCross = 5
  conUpwardDiagonal = 6
  conDownwardDiagonal = 7
  conHorizontal = 8
  conVertical = 9
End Enum

'Paint area resize direction constants declaration
Private Const conResizeWE = 0
Private Const conResizeNS = 1
Private Const conResizeNWSE = 2

'Default value
Private Const conDefaultActiveTool = conTPencil
Private Const conDefaultActiveFilterTool = conFltBrightness
Private Const conDefaultBorderStyle = vbBSSolid
Private Const conDefaultBrushShape = conFilledRect
Private Const conDefaultDotWidth = 0
Private Const conDefaultFillStyle = conTsBorderOnly
Private Const conDefaultInsideFillStyle = vbFSSolid
Private Const conDefaultPaintHeight = 6000
Private Const conDefaultPaintWidth = 6400

'Other constants declaration
Private Const conBufMax = 10               'maximum buffer for undo/redo feature
                                           '  (be careful increasing this value,
                                           '           it can make your computer
                                           '                  run out of memory)
Private Const conProgramTitle = "VB Paint"

'Variable Declaration
Private blnDrag As Boolean              'condition whether mouse move is to drag
Private blnDrawing As Boolean              'condition when mouse move is to draw
Private blnDrawingPolygon As Boolean                  'condition to draw polygon
Private blnFirstMoving As Boolean              'condition whether it's the first
                                               '   selected object moving action
Private blnMoving As Boolean                       'condition when mouse move is
                                                   ' to move the selected object
Private blnPicChanged As Boolean    'condition that the picture has been changed
                                    ' so the save confirmation on exit is needed
Private blnResize As Boolean      'condition that the paint area is being resize
Private lngDragStart As mdlAPI.typPoint  'coordinate where the drag action start
Private lngP1 As mdlAPI.typPoint         'the starting coordinate marked by user
Private lngP2 As mdlAPI.typPoint           'the ending coordinate marked by user
Private lngPolygon() As mdlAPI.typPoint     'to store polygon points information
Private intActiveFilterTool As enmFilter              'the active filter tool id
Private intActiveTool As enmTool     'the active tool id (active optTools index)
Private intBrushShape As enmBrushShape               'current active brush shape
Private intBufCur As Integer             'current buffer (for undo/redo feature)
Private intBufEnd As Integer           'last buffer used (for undo/redo feature)
Private intBufStart As Integer        'first buffer used (for undo/redo feature)
Private intDot As Integer                          'the width of the dot to draw
Private intDrawStyle As Integer                      'current .DrawSyle property
Private intFillStyle As enmFillStyle                'the current fill style used
Private intInsideFillStyle As Integer               'current .FillStyle property
Private sngZoomFactor As Single                             'current zoom factor
Private strFileName As String   'image file name (null string for unnamed image)


Private m_ColCommandButton              As Collection               ' 동적 생성 컨트롤 저장을 위한 컬렉션
Private WithEvents ClsEventMonitor      As ClassEventMonitor        ' 이벤트 전달을 위한 클래스
Attribute ClsEventMonitor.VB_VarHelpID = -1

'==== API 파일 오픈 관련 =================================================
Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 31
End Type
Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long          '  caller's window handle
        hDC As Long                '  printer DC/IC or NULL
        lpLogFont As Long          '  ptr. to a LOGFONT struct
        iPointSize As Long         '  10 * size in points of selected font
        flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESIZE or bigger
        nFontType As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type
Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Dim OFName As OPENFILENAME
Dim CustomColors() As Byte
'==== API 파일 오픈 관련 =================================================

' Purpose    : Adjust lngP2 coordinate to agree with Shift or Ctrl key status as
'              specified below:
'              - Shift key pressed is to draw a square shape like square,
'                circle, 45-degree line, etc.
'              - Ctrl key pressed and blnEnableCtrl = true are to draw a
'                horizontal or vertical shape, like horizontal line, vertical
'                line, etc.
' Assumption : These global variables has been initiated:
'                lngP1, lngP2
' Effect     : As specified
' Inputs     : * X (current X coordinate)
'              * Y (current Y coordinate)
'              * Shift (shift key status)
'              * blnEnableCtrl (condition whether ctrl key status will effect
'                               the drawing or not)
' Returns    : -
Private Sub AdjustP2(x As Single, y As Single, Shift As Integer, _
                     Optional blnEnableCtrl As Boolean = False)
  On Error GoTo ErrorHandler
  
  If Shift = vbShiftMask Then
    'Draw a square shape
    If Abs(x - lngP1.x) <= Abs(y - lngP1.y) Then
      lngP2.x = x
      If y > lngP1.y Then
        lngP2.y = lngP1.y + Abs(x - lngP1.x)
      Else
        lngP2.y = lngP1.y - Abs(x - lngP1.x)
      End If
    Else
      If x > lngP1.x Then
        lngP2.x = lngP1.x + Abs(y - lngP1.y)
      Else
        lngP2.x = lngP1.x - Abs(y - lngP1.y)
      End If
      lngP2.y = y
    End If
  ElseIf (Shift = vbCtrlMask) And blnEnableCtrl Then
    'Draw a horizontal or vertical shape
    If Abs(x - lngP1.x) <= Abs(y - lngP1.y) Then
      '- Horizontal shape
      lngP2.x = lngP1.x
      lngP2.y = y
    Else
      '- Vertical shape
      lngP2.x = x
      lngP2.y = lngP1.y
    End If
  Else
    'Draw a free shape
    lngP2.x = x
    lngP2.y = y
  End If
  Exit Sub
  
ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Adjust paint resize boxes (the little box on the right, bottom
'              abd bottom-right of the paint area) position to agree with paint
'              area width and height
' Assumption : These components exist in this form:
'                picPaint, picPaintResize
' Effect     : The paint resize boxes have been positioned to the middle right,
'              middle bottom and bottom-right next to the paint area
' Inputs     : -
' Returns    : -
Public Sub AdjustPaintResizeBox()
  On Error GoTo ErrorHandler
  
  picPaintResize(conResizeWE).Left = picPaint.Left + picPaint.Width
  picPaintResize(conResizeWE).Top = picPaint.Top + (picPaint.Height / 2)
  picPaintResize(conResizeNS).Left = picPaint.Left + (picPaint.Width / 2)
  picPaintResize(conResizeNS).Top = picPaint.Top + picPaint.Height
  picPaintResize(conResizeNWSE).Left = picPaintResize(conResizeWE).Left
  picPaintResize(conResizeNWSE).Top = picPaintResize(conResizeNS).Top
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Change the cursor in the paint area
' Assumptions: * This global variable has been initiated:
'                  intActiveTool
'              * This component exists in this form:
'                  picPaint
'              * The cursor file needed is exist in the sub directory "\Cursors"
' Effect     : The cursor in the paint area has been changed
' Inputs     : -
' Returns    : -
Private Sub ChangePaintCursor()
  On Error GoTo ErrorHandler                     'don't change the cursor if the
                                                 '     file needed doesn't exist
  With picPaint
    .MousePointer = vbCustom
    Select Case intActiveTool
      Case conTAirBrush
        .MouseIcon = LoadPicture(App.Path & "\Cursors\airbrush.cur")
      Case conTBrush
        .MouseIcon = LoadPicture(App.Path & "\Cursors\brush.cur")
      Case conTEraser
        .MouseIcon = LoadPicture(App.Path & "\Cursors\eraser.cur")
      Case conTFill
        .MouseIcon = LoadPicture(App.Path & "\Cursors\fill.cur")
      Case conTFilter
        .MouseIcon = LoadPicture(App.Path & "\Cursors\filter.cur")
      Case conTPencil
        .MouseIcon = LoadPicture(App.Path & "\Cursors\pencil.cur")
      Case conTPick
        .MouseIcon = LoadPicture(App.Path & "\Cursors\pick.cur")
      Case conTText
        .MouseIcon = LoadPicture(App.Path & "\Cursors\text.cur")
      Case conTSelect, conTCurve
        .MousePointer = vbDefault
      Case conTZoom
        .MouseIcon = LoadPicture(App.Path & "\Cursors\zoom.cur")
      Case conTHand
        .MouseIcon = LoadPicture(App.Path & "\Cursors\handflat.cur")
      Case Else
        .MouseIcon = LoadPicture(App.Path & "\Cursors\cross.cur")
    End Select
  End With

ErrorHandler:
End Sub

' Purpose    : Clear image buffer (for undo/redo feature)
' Assumption : These components exist in this form:
'                mnuRedo, mnuUndo, picBuffer(), picPaint
' Effects    : These global variables has been changed as following:
'              * intBufCur = 0
'              * intBufStart = 0
'              * intBufEnd = 0
'              * picBuffer.ubound = 0
'              * picBuffer(0).Picture = picPaint.Image
' Inputs     : -
' Returns    : -
Private Sub ClearImageBuffer()
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  intBufCur = 0
  intBufStart = 0
  intBufEnd = 0
  For i = 1 To picBuffer.UBound
    Unload picBuffer(i)
  Next
  picBuffer(intBufCur).Picture = picPaint.Image
  'save the paint area width and height for undo/redo action
  '  on resized paint area
  picBuffer(intBufCur).Tag = CStr((picPaint.Width * 100000) + picPaint.Height)
  mnuUndo.Enabled = False
  mnuRedo.Enabled = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw an air brush effect with current foreground color and
'              draw mode in the paint area
' Assumption : This component exist in this form:
'                picPaint
' Effects    : The air brush effect has been drawn in the paint area
' Inputs     : * X, Y (center coordinate of the air brush)
'              * R (half of the width or height of the air brush)
' Returns    : -
Private Sub DrawAirBrush(x As Integer, y As Integer, r As Integer)
  Const conIntencity = 0.25
  
  Dim i As Integer
  Dim intDrawWidth As Integer                  'to keep current draw width value
  
  On Error GoTo ErrorHandler
  
  With picPaint
    intDrawWidth = .DrawWidth
    .DrawWidth = 1
    Randomize
    For i = 1 To ((r * r) * conIntencity)
      picPaint.PSet (x - (r / 2) + (Rnd() * r), y - (r / 2) + (Rnd() * r))
    Next
    .DrawWidth = intDrawWidth
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw an arrow from (X1,Y1) to (X2,Y2) in the paint area with
'              current foreground color, draw mode and draw width in the paint
'              area
' Assumption : This component exists in this form:
'                picPaint
' Effect     : The arrow has been drawn in the paint area
' Inputs     : X1, Y1, X2, Y2
' Returns    : -
Private Sub DrawArrow(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
  Const conAlphaTip = 45                  'the angle of the lines in arrow's tip
  Const conLenTip = 10                       'length of the lines in arrow's tip
  Const conPi = 3.14159265358979
  
  'Variables to draw arrow's tip
  Dim intSign As Integer
  Dim X3 As Integer
  Dim Y3 As Integer
  Dim X4 As Integer
  Dim Y4 As Integer
  Dim sngBeta As Single
  
  On Error GoTo ErrorHandler
  
  'Draw arrow's line
  picPaint.Line (X1, Y1)-(X2, Y2)
  'Calculate variables for arrow's tip
  If X2 - X1 <> 0 Then
    sngBeta = Atn((Y2 - Y1) / (X2 - X1)) * 180 / conPi
  Else
    sngBeta = 90
  End If
  If X2 > X1 Then
    intSign = 1
  ElseIf X2 < X1 Then
    intSign = -1
  ElseIf Y2 > Y1 Then
    intSign = 1
  ElseIf Y2 < Y1 Then
    intSign = -1
  End If
  X3 = X2 - ((conLenTip * Cos((conAlphaTip + sngBeta) * conPi / 180)) * intSign)
  Y3 = Y2 - ((conLenTip * Sin((conAlphaTip + sngBeta) * conPi / 180)) * intSign)
  X4 = X2 - ((conLenTip * Cos((conAlphaTip - sngBeta) * conPi / 180)) * intSign)
  Y4 = Y2 + ((conLenTip * Sin((conAlphaTip - sngBeta) * conPi / 180)) * intSign)
  'Draw arrow's tip
  picPaint.Line (X2, Y2)-(X3, Y3)
  picPaint.Line (X2, Y2)-(X4, Y4)
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw a brush shape intBrushShape at (X,Y) with
'              current foreground color, draw style and fill style in the paint
'              area
' Assumption : This component exists in this form:
'                picPaint
' Effect     : The brush shape has been drawn in the paint area
' Inputs     : intBrushShape, X, Y
' Returns    : -
Private Sub DrawBrush(intBrushShape As enmBrushShape, x As Single, y As Single)
  Const conBrushSize = 3
  
  Dim intDrawWidth As Integer       'to keep current picPaint.DrawWidth property
  
  On Error GoTo ErrorHandler
  
  With picPaint
    intDrawWidth = .DrawWidth
    .DrawWidth = 1
    Select Case intBrushShape
      Case conFilledRect
        picPaint.FillStyle = intInsideFillStyle
        picPaint.Line (x - (conBrushSize * intDrawWidth), _
                       y - (conBrushSize * intDrawWidth))- _
                      (x + (conBrushSize * intDrawWidth), _
                       y + (conBrushSize * intDrawWidth)), , BF
      Case conFilledCircle
        picPaint.FillStyle = intInsideFillStyle
        picPaint.Circle (x, y), conBrushSize * intDrawWidth
      Case conRect
        picPaint.FillStyle = vbFSTransparent
        picPaint.Line (x - (conBrushSize * intDrawWidth), _
                       y - (conBrushSize * intDrawWidth))- _
                      (x + (conBrushSize * intDrawWidth), _
                       y + (conBrushSize * intDrawWidth)), , B
      Case conCircle
        picPaint.FillStyle = vbFSTransparent
        picPaint.Circle (x, y), conBrushSize * intDrawWidth
      Case conCross
        picPaint.Line (x - (conBrushSize * intDrawWidth), y)- _
                      (x + (conBrushSize * intDrawWidth), y)
        picPaint.Line (x, y - (conBrushSize * intDrawWidth))- _
                      (x, y + (conBrushSize * intDrawWidth))
      Case conDiagonalCross
        picPaint.Line (x - (conBrushSize * intDrawWidth), _
                       y + (conBrushSize * intDrawWidth))- _
                      (x + (conBrushSize * intDrawWidth), _
                       y - (conBrushSize * intDrawWidth))
        picPaint.Line (x - (conBrushSize * intDrawWidth), _
                       y - (conBrushSize * intDrawWidth))- _
                      (x + (conBrushSize * intDrawWidth), _
                       y + (conBrushSize * intDrawWidth))
      Case conUpwardDiagonal
        picPaint.Line (x - (conBrushSize * intDrawWidth), _
                       y + (conBrushSize * intDrawWidth))- _
                      (x + (conBrushSize * intDrawWidth), _
                       y - (conBrushSize * intDrawWidth))
      Case conDownwardDiagonal
        picPaint.Line (x - (conBrushSize * intDrawWidth), _
                       y - (conBrushSize * intDrawWidth))- _
                      (x + (conBrushSize * intDrawWidth), _
                       y + (conBrushSize * intDrawWidth))
      Case conHorizontal
        picPaint.Line (x - (conBrushSize * intDrawWidth), y)- _
                      (x + (conBrushSize * intDrawWidth), y)
      Case conVertical
        picPaint.Line (x, y - (conBrushSize * intDrawWidth))- _
                      (x, y + (conBrushSize * intDrawWidth))
    End Select
    .DrawWidth = intDrawWidth
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw a bezier curve in the paint area with current foreground
'              color, draw mode (except for blnComplete = true which is
'              draw mode = copy) and draw width
' Assumption : These components exist in this form:
'                picPaint, imgBezier()
' Effect     : The bezier curve has been drawn in the paint area
' Inputs     : * blnCreate (condition to "create" [not to edit] the curve)
'              * blnComplete (condition to finish [draw with copy mode] the
'                             curve drawing)
'              * X, Y (center coordinate of the curve)
' Returns    : -
Private Sub DrawCurveBezier(Optional blnCreate As Boolean = False, _
                            Optional blnComplete As Boolean = False, _
                            Optional x As Single, Optional y As Single)
  Const conCreateRadius = 50
  
  Dim i As Integer
  Dim intScaleMode                             'to keep current scale mode value
  Dim lngBezier(3) As typPoint
  
  On Error GoTo ErrorHandler
  
  intScaleMode = picPaint.ScaleMode
  picPaint.ScaleMode = vbPixels
  If blnCreate Then
    imgBezier(0).Top = y - conCreateRadius
    imgBezier(0).Left = x - conCreateRadius
    imgBezier(1).Top = y - conCreateRadius
    imgBezier(1).Left = x + conCreateRadius
    imgBezier(2).Top = y + conCreateRadius
    imgBezier(2).Left = x - conCreateRadius
    imgBezier(3).Top = y + conCreateRadius
    imgBezier(3).Left = x + conCreateRadius
    For i = 0 To 3
      imgBezier(i).Visible = True
    Next
  End If
  lngBezier(0).x = imgBezier(0).Left + (imgBezier(0).Width / 2)
  lngBezier(0).y = imgBezier(0).Top + (imgBezier(0).Height / 2)
  lngBezier(1).x = imgBezier(1).Left + (imgBezier(0).Width / 2)
  lngBezier(1).y = imgBezier(1).Top + (imgBezier(0).Height / 2)
  lngBezier(2).x = imgBezier(2).Left + (imgBezier(0).Width / 2)
  lngBezier(2).y = imgBezier(2).Top + (imgBezier(0).Height / 2)
  lngBezier(3).x = imgBezier(3).Left + (imgBezier(0).Width / 2)
  lngBezier(3).y = imgBezier(3).Top + (imgBezier(0).Height / 2)
  With picPaint
    If blnComplete Then
      .DrawMode = vbCopyPen
    End If
    mdlAPI.PolyBezier picPaint.hDC, lngBezier(0), 4
    .Refresh
  End With
  picPaint.ScaleMode = intScaleMode
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw polygon lngPolygon() in the paint area with current
'              foreground color, draw mode (except for blnComplete = true which
'              is draw mode = copy and foreground color depends on fill style)
'              and draw width (also with fill style intFillStyle and fill color
'              lblFillColor.BackColor for blnComplete = true)
' Assumption : * These components exist in this form:
'                  picPaint, imgBezier(), lnlForeColor, lblFillColor
'              * These global variables has been initiated
'                  lngPolygon()
' Effect     : The polygon has been drawn in the paint area
' Inputs     : * blnComplete (condition to finsih [draw with copy mode and fill
'                             style intFillStyle)
'              * blnOnlyDrawLastLine (condition to draw only the last line of
'                                     the polygon)
' Returns    : -
Private Sub DrawPolygon(Optional blnComplete As Boolean = True, _
                        Optional blnOnlyDrawLastLine = True)
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  With picPaint
    If blnComplete Then
      .DrawMode = vbCopyPen
      Select Case intFillStyle
        Case conTsBorderOnly
          .FillStyle = vbFSTransparent
          .ForeColor = lblForeColor.BackColor
        Case conTsBorderFill
          .FillStyle = intInsideFillStyle
          .ForeColor = lblForeColor.BackColor
          .FillColor = lblFillColor.BackColor
        Case conTsFillOnly
          .FillStyle = intInsideFillStyle
          .ForeColor = lblFillColor.BackColor
          .FillColor = lblFillColor.BackColor
      End Select
      mdlAPI.Polygon picPaint.hDC, lngPolygon(0), UBound(lngPolygon) + 1
      .Refresh
    Else
      If UBound(lngPolygon) > 0 Then
        If blnOnlyDrawLastLine Then
          picPaint.Line (lngPolygon(UBound(lngPolygon) - 1).x, _
                         lngPolygon(UBound(lngPolygon) - 1).y)- _
                        (lngPolygon(UBound(lngPolygon)).x, _
                         lngPolygon(UBound(lngPolygon)).y)
        Else
          For i = 1 To UBound(lngPolygon)
            picPaint.Line (lngPolygon(i - 1).x, lngPolygon(i - 1).y)- _
                          (lngPolygon(i).x, lngPolygon(i).y)
          Next
        End If
      End If
    End If
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Draw selection rectangle with xor mode and dot style
' Assumption : These components exist in this form:
'                picPaint, picSelect
' Effect     : As specified
' Inputs     : -
' Returns    : -
Public Sub DrawSelectionRect()
  'Variables to keep picPaint properties
  Dim intDrawStyle As Integer
  Dim intDrawMode As Integer
  Dim intDrawWidth As Integer
  
  On Error GoTo ErrorHandler
  
  If picSelect.Visible Then
    With picPaint
      intDrawMode = .DrawMode
      intDrawWidth = .DrawWidth
      picPaint.DrawStyle = vbDot
      picPaint.DrawMode = vbXorPen
      picPaint.DrawWidth = 1
      blnFirstMoving = False
      picPaint.Line (picSelect.Left - 1, picSelect.Top - 1)- _
                    (picSelect.Left + picSelect.Width, _
                     picSelect.Top + picSelect.Height), _
                    vbBlack Xor picPaint.BackColor, B
      .DrawStyle = intDrawStyle
      .DrawMode = intDrawMode
      .DrawWidth = intDrawWidth
    End With
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub


'-- 화면 초기화
Private Sub FrmInitial()
    Dim x As Printer
    Dim prtSelectPrinter As Printer
    Dim boolPrinter_Select_Fales As Boolean
    Dim strDefault As String
    Dim Buffer As String
    Dim aryPrinter() As String
    Dim strBuffer As String
    Dim i As Integer
    Dim j As Integer
 
    ' 클래스 초기화
    Set ClsEventMonitor = New ClassEventMonitor
    Set m_ColCommandButton = New Collection

    Call CtrlInitial
    
    '구분
    cboType.Clear
    cboType.AddItem "S_Text"
    cboType.AddItem "D_Text"
    cboType.AddItem "S_Image"
    cboType.AddItem "D_Image"
    cboType.AddItem "Barcode"
    cboType.AddItem "Line"
    
    cboType.ListIndex = 0
    
    '바코드 타입
    cboBarType.Clear
    cboBarType.AddItem "None"
    cboBarType.AddItem "2of5[공통]"               '5
    cboBarType.AddItem "Interleaved2of5[공통]"    '6
    cboBarType.AddItem "3of9[공통]"               '0
    cboBarType.AddItem "Codabar[공통]"            '9
    cboBarType.AddItem "3of9X[공통]"              '1
    cboBarType.AddItem "Code128A[공통]"           '11
    cboBarType.AddItem "Code128B[공통]"           '12
    cboBarType.AddItem "Code128C[공통]"           '13
    cboBarType.AddItem "UPCA[공통]"               '15
    cboBarType.AddItem "MSI[공통]"                '7
    cboBarType.AddItem "Code93[공통]"             '3
    cboBarType.AddItem "ExtendedCode93[공통]"     '4
    cboBarType.AddItem "EAN13[공통]"              '17
    cboBarType.AddItem "EAN8[공통]"               '18
    cboBarType.AddItem "PostNet[공통]"            '23
    cboBarType.AddItem "ANSI3of9[신규]"           '
    cboBarType.AddItem "ANSI3of9X[신규]"          '
    cboBarType.AddItem "Code128Auto[공통]"        '10
    cboBarType.AddItem "UCCEAN128[공통]"          '27
    cboBarType.AddItem "UPCE[공통]"               '16
    cboBarType.AddItem "RoyalMail[신규]"          '
    cboBarType.AddItem "MSICode2[공통]"           '8  ??MSIPlessey
    cboBarType.AddItem "DUN14[공통]"              '28
    
    cboBarType.ListIndex = 7
    
' 0:Code39
' 1:Code39Extended
' 2:Code39Trioptic  x
' 3:Code93
' 4:Code93Extended
' 5:Code2of5
' 6:Interleave2of5
' 7:MSICode
' 8:MSIPlessey
' 9:Codabar
'10:Code128
'11:Code128A
'12:Code128B
'13:Code128C
'14:Code11          x
'15:UPCA
'16:UPCE
'17:EAN13
'18:EAN8
'19:EAN99           x
'20:JAN8            x
'21:JAN13           x
'22:Telepen         x
'23:PostNet
'24:RM4SCC          x
'25:PZN             x
'26:ISBN            x
'27:UCCEAN128       x
'28:DUN14           x
    
    
    With spdList
        .MaxRows = 0
        .MaxCols = 29
'        .SetText 1, 0, "설정순번"
'        .SetText 2, 0, "항목구분"
'        .SetText 3, 0, "항목명"
'        .SetText 4, 0, "X1좌표"
'        .SetText 5, 0, "X2좌표"
'        .SetText 6, 0, "Y1좌표"
'        .SetText 7, 0, "Y2좌표"
'        .SetText 8, 0, "폰트명"
'        .SetText 9, 0, "폰트사이즈"
'        .SetText 10, 0, "굵기"
'        .SetText 11, 0, "비틀림"
'        .SetText 12, 0, "밑줄"
'        .SetText 13, 0, "폰트회전"
'        .SetText 14, 0, "바코드종류"
'        .SetText 15, 0, "바코드폭"
'        .SetText 16, 0, "바코드회전"
'        .SetText 17, 0, "이미지경로"
'        .SetText 18, 0, "라인회전"
'        .SetText 19, 0, "라인두께"
'        .SetText 20, 0, "라인폭"
'        .SetText 21, 0, "출력여부"
'        .SetText 22, 0, "출력값"
'        .SetText 23, 0, "X좌표 보정값"
'        .SetText 24, 0, "Y좌표 보정값"
'        .SetText 25, 0, "용지높이"
'        .SetText 26, 0, "용지폭"
'        .SetText 27, 0, "무조건고정"
'        .SetText 28, 0, "용지방향"
'        .SetText 29, 0, "Tag"
'        .ColWidth(-1) = 10 '10
'        .ColWidth(29) = 0
    End With
    
    '-- 프린터
    For Each x In Printers
        cmbPrinter.AddItem x.DeviceName
    Next
    
    strBuffer = Space(1024)
 
    i = GetProfileString("windows", "Device", "", strBuffer, Len(strBuffer))
    aryPrinter = Split(strBuffer, ",")
    strDefault = Trim(aryPrinter(0))
 
    For Each prtSelectPrinter In Printers
        j = j + 1
        If UCase(Trim(prtSelectPrinter.DeviceName)) = UCase(Trim(strDefault)) Then
            Set Printer = prtSelectPrinter
            boolPrinter_Select_Fales = True
            cmbPrinter.ListIndex = j - 1
            Exit For
        End If
    Next
    
    '-- 가로
    If optHW(0).Value = True Then
        txtPaperHSize.Text = ""
        txtPaperWSize.Text = ""
        
    '-- 세로
    Else
    
    End If
    
    '-- Mode Set
    intMode = 0

    '-- 바코드 이미지명 초기화
    strBarImgName = ""
    
End Sub

Private Sub Command2_Click()
    Dim strSrcfile  As Variant
    Dim varBuffer() As Variant
    Dim varBuf      As Variant
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim j           As Long
    Dim bytBuff()   As Byte
    
    Static ChkSumCnt As Long
    Dim strTxt As String
    
    Dim FileNumber As Long
    Dim FileName As String
    Dim FileCount As Long
    Dim LineCount As Long
    Dim FileOpenNumber As Integer
    Dim data As String
    Dim splitdata() As String
    
    Dim utf8() As Byte
    Dim ucs2 As Variant
    Dim chars As Long
    Dim varTmp As Variant
    
    Me.ScaleMode = gScaleMode
    
    ' 폼초기화
    Call FrmInitial
    
    'Cancel을 True로 설정합니다.
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
     
    '경로 속성을 설정합니다.
    CommonDialog1.InitDir = App.Path & "\" & gLayOut
    CommonDialog1.Filter = "LayoutFile(*.lof)|*.lof"
    
    '[파일] 대화 상자를 표시합니다.
    CommonDialog1.ShowOpen
    strSrcfile = CommonDialog1.FileName

    '컬렉션 초기화
    Set m_ColCommandButton = Nothing
    Set m_ColCommandButton = New Collection
    
    'LOF 파일 열기
    FileName = CommonDialog1.FileName
    varTmp = Split(FileName, "\")
    Me.Caption = varTmp(UBound(varTmp))
    FileOpenNumber = FreeFile()
    LineCount = 0
    
    Open FileName For Binary As #1   'UTF-8 문서지정
    ReDim utf8(LOF(1))
    
    Get #1, , utf8
    
    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), 0, 0)
    ucs2 = Space(chars)
    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), StrPtr(ucs2), chars)
    varBuf = Split(ucs2, Chr(13))
    
    Close #1
    
    
    '오픈한 LOF파일 버퍼에 쓰기
    For i = 0 To UBound(varBuf)
        ReDim Preserve varBuffer(i)
        varBuffer(LineCount) = varBuf(i)
        LineCount = LineCount + 1
    Next
            
    '오픈한 LOF파일 화면그리기/스프레드쓰기
    For i = 0 To UBound(varBuffer) - 1
        If varBuffer(i) <> "" Then
            varBuf = Split(varBuffer(i), "^")
            Call MakeLayout(varBuf)
            Call SetList(varBuf)
        End If
    Next
    
'    txtText_DblClick
'    SetImageBuffer
    
'    intMode = 1
    
    Exit Sub

ErrHandler:
End Sub

Private Sub Form_Activate()
  On Error GoTo ErrorHandler
  
  picPaint.SetFocus
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler
' If Now() > "2009-07-9" Then
'    MsgBox "데모버전입니다.", vbCritical, "[데모!]"
'    End
' End If
  mnuNew_Click
  'Init default value
  intActiveFilterTool = conDefaultActiveFilterTool
  intActiveTool = conDefaultActiveTool
  intBrushShape = conDefaultBrushShape
  intDot = conDefaultDotWidth
  intInsideFillStyle = conDefaultFillStyle
  intFillStyle = conDefaultFillStyle
  mnuFilterTools(intActiveFilterTool).Checked = True
  picPaint.BorderStyle = conDefaultBorderStyle
  'Init dialogs' flags
  cdlSave.flags = cdlOFNHideReadOnly Or _
                  cdlOFNOverwritePrompt Or cdlOFNPathMustExist
  cdlOpen.flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
  cdlFonts.flags = cdlCFBoth Or cdlCFEffects Or cdlCFForceFontExist
  cdlPrint.flags = cdlPDNoPageNums Or cdlPDNoSelection
  'Init fonts
  With picPaint
    .FontBold = txtText.FontBold
    .FontItalic = txtText.FontItalic
    .FontName = txtText.FontName
    .FontSize = txtText.FontSize
    .FontStrikethru = txtText.FontStrikethru
    .FontUnderline = txtText.FontUnderline
  End With
  'Init paint area size
  picPaint.Width = conDefaultPaintWidth
  picPaint.Height = conDefaultPaintHeight
  AdjustPaintResizeBox
  'Others
  UpdateStatusBar
  ChangePaintCursor
  
  
  
  
    Dim x As Printer
    Dim prtSelectPrinter As Printer
    Dim boolPrinter_Select_Fales As Boolean
    Dim strDefault As String
    Dim Buffer As String
    Dim aryPrinter() As String
    Dim strBuffer As String
    Dim i As Integer
    Dim j As Integer
'    Dim strLicense As String
'    Dim strKey  As String
'
'    strLicense = "License"
'
'    strKey = GetString(HKEY_CURRENT_USER, REG_POSITION, strLicense)
'
'    If strKey = "" Or Not IsDate(strKey) And strKey < Format(Now) Then
'        MsgBox "라이센스 기간이 만료되었거나 없습니다." & vbNewLine & "개발자에게 문의하십시요", vbCritical + vbOKOnly, Me.Caption
'        End
'    End If
        
    ' 버전 정보 표시
    Me.Caption = Me.Caption & " [Ver " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    'Combo1.ListIndex = 1
    
    Call MDIForm_Tool
    
    Call FrmInitial

    Call GetSetup
        
    txtDevide.Text = gDevide
    
    
    '==== API 파일 오픈 관련 =================================================
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    
    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 0
    Next i
    '==== API 파일 오픈 관련 =================================================
    
    Me.Top = 0
    Me.Left = 0
'    Me.ScaleMode = gScaleMode
    
'    Picture1.ScaleMode = vbTwips
      
  
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Public Sub Form_Resize()
  On Error GoTo ErrorHandler
  
  If Me.WindowState <> vbMinimized Then
    'Limit the form's height
    If Me.Height < 2800 Then
      Me.Height = 2800
    End If
    'Adjust the tools and color box position and size
    fraTools.Height = Me.ScaleHeight - 900
    fraColor.Top = Me.ScaleHeight - 1110
    fraColor.Width = Me.Width - 90
    'Adjust the vertical scroll bar position, size and other properties
    With vscPaint
      If hscPaint.Visible Then
        .Max = (picPaint.Height - (Me.Height - hscPaint.Height - 1950)) / 10
      Else
        .Max = (picPaint.Height - (Me.Height - 1950)) / 10
      End If
      .Visible = (.Max > 0)
      If .Visible Then
        .Left = Me.Width - .Width - 110
        If hscPaint.Visible Then
          .Height = Me.ScaleHeight - fraColor.Height - hscPaint.Height - 150
        Else
          .Height = Me.ScaleHeight - fraColor.Height - 150
        End If
      End If
    End With
    'Adjust the horizontal scroll bar position, size and other properties
    With hscPaint
      If vscPaint.Visible Then
        .Max = (picPaint.Width - (Me.Width - vscPaint.Width - 1050)) / 10
      Else
        .Max = (picPaint.Width - (Me.Width - 1050)) / 10
      End If
      .Visible = (.Max > 0)
      If .Visible Then
        .Top = fraColor.Top - .Height + 110
        If vscPaint.Visible Then
          .Width = Me.Width - fraTools.Width - vscPaint.Width - 90
        Else
          .Width = Me.Width - fraTools.Width - 90
        End If
      End If
    End With
    'Re-adjust the vertical scroll bar max and height to match the new
    '  horizontal scroll bar properties
    If hscPaint.Visible Then
      vscPaint.Max = (picPaint.Height - _
                      (Me.Height - hscPaint.Height - 1850)) / 10
      vscPaint.Height = Me.ScaleHeight - fraColor.Height - hscPaint.Height - 150
    End If
    'Adjust the fraScroll properties
    If hscPaint.Visible And vscPaint.Visible Then
      fraScroll.Visible = True
      fraScroll.Left = vscPaint.Left
      fraScroll.Top = hscPaint.Top
    Else
      fraScroll.Visible = False
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ErrorHandler
  
  Dim intSave As Integer

  'Save confirmation
  If blnPicChanged = True Then
    intSave = MsgBox("Do you want to save the changes?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intSave
      Case vbYes
        mnuSave_Click
        Cancel = blnPicChanged
      Case vbCancel
        Cancel = True
    End Select
  End If
  
    ' 컬렉션 초기화
    Set m_ColCommandButton = Nothing
    Set ClsEventMonitor = Nothing
  
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub fraOptDot_MouseDown(Button As Integer, _
                                Shift As Integer, x As Single, y As Single)
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  For i = 0 To 7
    shpDot(i).FillColor = vbBlack
    shpDot(i).BorderColor = vbBlack
  Next
  'Set draw width intDot value and highlight the tool based on mouse click
  '  coordinate (X,Y)
  If Button = vbLeftButton Then
    If (y >= 150) And (y < 400) Then
      lblDot.Top = 150
      If (x >= 75) And (x < 325) Then
        intDot = 0
        lblDot.Left = 75
      ElseIf (x >= 325) And (x < 575) Then
        intDot = 1
        lblDot.Left = 325
      End If
    ElseIf (y >= 400) And (y < 650) Then
      lblDot.Top = 400
      If (x >= 75) And (x < 325) Then
        intDot = 2
        lblDot.Left = 75
      ElseIf (x >= 325) And (x < 575) Then
        intDot = 3
        lblDot.Left = 325
      End If
    ElseIf (y >= 650) And (y < 900) Then
      lblDot.Top = 650
      If (x >= 75) And (x < 325) Then
        shpDot(4).FillColor = vbWhite
        intDot = 4
        lblDot.Left = 75
      ElseIf (x >= 325) And (x < 575) Then
        intDot = 5
        lblDot.Left = 325
      End If
    ElseIf (y >= 900) And (y < 1150) Then
      lblDot.Top = 900
      If (x >= 75) And (x < 325) Then
        intDot = 6
        lblDot.Left = 75
      ElseIf (x >= 325) And (x < 575) Then
        intDot = 7
        lblDot.Left = 325
      End If
    End If
    shpDot(intDot).FillColor = vbWhite
    shpDot(intDot).BorderColor = vbWhite
    'Update the current drawing to match the new draw width
    UpdateDrawing
    picPaint.DrawWidth = intDot + 1
    UpdateDrawing
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub fraOptFill_MouseDown(Button As Integer, _
                                 Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  'Set fill style intFillStyle and highlight the tool based on mouse click
  '  coordinate (X,Y)
  If Button = vbLeftButton Then
    If (y >= 125) And (y < 425) Then
      shpRect(0).BorderColor = vbWhite
      shpRect(1).BorderColor = vbBlack
      shpRect(2).BorderColor = vbBlack
      lblFill.Top = 150
      intFillStyle = conTsBorderOnly
    ElseIf (y >= 450 And y < 750) Then
      shpRect(0).BorderColor = vbBlack
      shpRect(1).BorderColor = vbWhite
      shpRect(2).BorderColor = vbBlack
      lblFill.Top = 465
      intFillStyle = conTsBorderFill
    ElseIf (y >= 775 And y < 1075) Then
      shpRect(0).BorderColor = vbBlack
      shpRect(1).BorderColor = vbBlack
      shpRect(2).BorderColor = vbWhite
      lblFill.Top = 780
      intFillStyle = conTsFillOnly
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub hscPaint_Change()
  Dim lngPicPaintLeft As Long
  
  On Error GoTo ErrorHandler
  
  lngPicPaintLeft = CLng(fraTools.Width) - (CLng(hscPaint.Value) * 10)
  picPaint.Left = lngPicPaintLeft
  AdjustPaintResizeBox
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Apply effect intEffect to the selection (if any) or to the paint
'              area
' Assumption : These components exist in this form:
'                mnuRotateClockWise, picPaint, picSelection, picImageEffect
' Effect     : As specified
' Inputs     : intImageEffect, sngResizeFactor
' Returns    : -
Private Sub ImageEffect(intEffect As enmEffect, _
                        Optional sngResizeFactor As Single, _
                        Optional sngRotateAngle As Single)
  Dim pic As PictureBox
  
  On Error GoTo ErrorHandler

  If picSelect.Visible Then
    Set pic = picSelect
  Else
    picPaint_DblClick
    Set pic = picPaint
  End If
  Select Case intEffect
    Case conEffResize
      If Not mnuResizeHeight.Checked Then
        mdlEffect.sngResizeWidth = sngResizeFactor
      End If
      If Not mnuResizeWidth.Checked Then
        mdlEffect.sngResizeHeight = sngResizeFactor
      End If
    Case conEffRotate
      mdlEffect.blnRotateClockWise = mnuRotateClockwise.Checked
      mdlEffect.sngRotateAngle = sngRotateAngle
  End Select
  If (intEffect <> conEffResize) Or _
     ((pic.ScaleWidth * Screen.TwipsPerPixelX * sngResizeFactor <= _
       mdlEffect.conMaxImageWidth) And _
      (pic.ScaleHeight * Screen.TwipsPerPixelY * sngResizeFactor <= _
       mdlEffect.conMaxImageHeight)) Then
    mdlEffect.ApplyEffect intEffect:=intEffect, _
                          pic:=pic, picTemp:=picImageEffect
  End If
  DrawSelectionRect
  If Not picSelect.Visible Then
    SetImageBuffer
  End If
  DrawSelectionRect
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Apply filter intFilter to the selection (if any) or to the paint
'              area
' Assumption : These components exist in this form:
'                picPaint, picSelection, lblForeColor, lblFillColor
' Effect     : As specified
' Input      : intFilter
' Returns    : -
Private Sub ImageFilter(intFilter As enmFilter, _
                        Optional x As Long = -1, Optional y As Long = -1)
  On Error GoTo ErrorHandler
  
  Dim pic As PictureBox
  Dim X1 As Long
  Dim Y1 As Long
  Dim X2 As Long
  Dim Y2 As Long
  
  If picSelect.Visible Then
    Set pic = picSelect
  Else
    picPaint_DblClick
    Set pic = picPaint
  End If
  If intFilter = conFltReplaceColors Then
    mdlFilter.lngReplacedColor = lblForeColor.BackColor
    mdlFilter.lngReplaceWithColor = lblFillColor.BackColor
  End If
  If (intActiveTool = conTFilter) And ((x <> -1) Or (y <> -1)) Then
    X1 = x - intDot
    Y1 = y - intDot
    X2 = x + intDot
    Y2 = y + intDot
    If (X2 >= 0) And (Y2 >= 0) Then
      mdlFilter.ApplyFilter intFilter:=intFilter, pic:=picPaint, _
                            X1:=X1, Y1:=Y1, X2:=X2, Y2:=Y2
    End If
  Else
    mdlFilter.ApplyFilter intFilter:=intFilter, pic:=pic
    DrawSelectionRect
    If Not picSelect.Visible Then
      SetImageBuffer
    End If
    DrawSelectionRect
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Zoom the paint area sngZoomFactor times (or no zoom if
'              blnNoZoom = true) and adjust the scroll bar so the coordinate
'              clicked by users is positioned in the center of visible
'              paint area
' Assumption : These components exist in this form:
'                picPaint, picZoom, picImageEffect, picPaintResize, hscPaint,
'                vcsPaint
' Effect     : * As specified
'              * if blnNoZoom = true then picPaintResize is shown, else
'                picPaintResize is hidden
' Inputs     : * X, Y (coordinate (in pixel) clicked by users that has been
'                      adjusted with zoom factor)
'              * blnNoZoom
' Returns    : -
Private Sub ImageZoom(Optional x As Long = 0, Optional y As Long = 0, _
                      Optional blnNoZoom As Boolean = False)
  Dim lngHscValue As Long                  'adjusted horizontal scroll bar value
  Dim lngVscValue As Long                    'adjusted vertical scroll bar value
  Dim lngVisibleWidth As Long                   'the width of visible paint area
  Dim lngVisibleHeight As Long                 'the height of visible paint area
  
  On Error GoTo ErrorHandler
  
  If blnNoZoom Then
    If sngZoomFactor <> 1 Then
      sngZoomFactor = 1
      picPaint.Picture = picZoom.Image
      frmPaint.AdjustPaintResizeBox
      frmPaint.Form_Resize
      picPaintResize(0).Visible = True
      picPaintResize(1).Visible = True
      picPaintResize(2).Visible = True
    End If
  Else
    'Zoom the picture
    mdlEffect.sngResizeWidth = sngZoomFactor
    mdlEffect.sngResizeHeight = sngZoomFactor
    picPaintResize(0).Visible = False
    picPaintResize(1).Visible = False
    picPaintResize(2).Visible = False
    picPaint.Visible = False
    picPaint.Picture = picZoom.Image
    mdlEffect.ApplyEffect intEffect:=conEffResize, _
                          pic:=picPaint, picTemp:=picImageEffect
    'Arrange horizontal scroll bar value
    If hscPaint.Visible Then
      If vscPaint.Visible Then
        lngVisibleWidth = Me.Width - fraTools.Width - vscPaint.Width
      Else
        lngVisibleWidth = Me.Width - fraTools.Width
      End If
      lngHscValue = ((x - (lngVisibleWidth / 2)) / _
                     (picPaint.Width - lngVisibleWidth)) * hscPaint.Max
      If lngHscValue < 0 Then
        hscPaint.Value = 0
      ElseIf lngHscValue > hscPaint.Max Then
        hscPaint.Value = hscPaint.Max
      Else
        hscPaint.Value = lngHscValue
      End If
    End If
    'Arrange vertical scroll bar value
    If vscPaint.Visible Then
      If hscPaint.Visible Then
        lngVisibleHeight = Me.ScaleHeight - _
                           hscPaint.Height - fraColor.Height - sta.Height
      Else
        lngVisibleHeight = Me.ScaleHeight - fraColor.Height - sta.Height
      End If
      lngVscValue = ((y - (lngVisibleHeight / 2)) / _
                     (picPaint.Height - lngVisibleHeight)) * vscPaint.Max
      If lngVscValue < 0 Then
        vscPaint.Value = 0
      ElseIf lngVscValue > vscPaint.Max Then
        vscPaint.Value = vscPaint.Max
      Else
        vscPaint.Value = lngVscValue
      End If
    End If
    picPaint.SetFocus
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub imgBezier_MouseDown(Index As Integer, Button As Integer, _
                                Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  'Start the drag operation on imgBezier(Index)
  lngDragStart.x = CLng(x)
  lngDragStart.y = CLng(y)
  blnDrag = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub imgBezier_MouseMove(Index As Integer, Button As Integer, _
                                Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  'Move imgBezier(Index) for the drag operation and update the bezier curve
  If blnDrag Then
    DrawCurveBezier
    picPaint.ScaleMode = vbTwips
    With imgBezier(Index)
      .Top = .Top + (y - lngDragStart.y)
      .Left = .Left + (x - lngDragStart.x)
    End With
    DrawCurveBezier
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub imgBezier_MouseUp(Index As Integer, Button As Integer, _
                              Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  'End the drag operation on imgBezier(Index)
  blnDrag = False
  picPaint.ScaleMode = vbPixels
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub imgBrush_MouseDown(Index As Integer, Button As Integer, _
                               Shift As Integer, x As Single, y As Single)
 On Error GoTo ErrorHandler
  
  intBrushShape = Index
  lblBrush.Top = imgBrush(Index).Top - (4 * Screen.TwipsPerPixelX)
  lblBrush.Left = imgBrush(Index).Left - (4 * Screen.TwipsPerPixelY)
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub lblColor_MouseDown(Index As Integer, Button As Integer, _
                               Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  Select Case Button
    Case vbLeftButton
      'Set the foreground color and update the current drawing to match the new
      '  foreground color
      UpdateDrawing
      lblForeColor.BackColor = lblColor(Index).BackColor
      picPaint.DrawMode = vbXorPen
      picPaint.ForeColor = picPaint.BackColor Xor lblForeColor.BackColor
      UpdateDrawing
    Case vbRightButton
      'Set the background color
      lblFillColor.BackColor = lblColor(Index).BackColor
  End Select
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub lblColor_MouseMove(Index As Integer, Button As Integer, _
                               Shift As Integer, x As Single, y As Single)
  UpdateStatusBar intInfo:=conStColorBox
End Sub

Private Sub lblFillColor_DblClick()
  On Error GoTo ErrorHandler
  
  cdlColor.ShowColor
  lblFillColor.BackColor = cdlColor.Color
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub lblFillColor_MouseMove(Button As Integer, _
                                   Shift As Integer, x As Single, y As Single)
  UpdateStatusBar intInfo:=conStBackColorBox
End Sub

Private Sub lblForeColor_DblClick()
  On Error GoTo ErrorHandler
  
  cdlColor.ShowColor
  'Update the current drawing to match with the new foreground color
  UpdateDrawing
  lblForeColor.BackColor = cdlColor.Color
  picPaint.DrawMode = vbXorPen
  picPaint.ForeColor = picPaint.BackColor Xor lblForeColor.BackColor
  UpdateDrawing
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub lblForeColor_MouseMove(Button As Integer, _
                                   Shift As Integer, x As Single, y As Single)
  UpdateStatusBar intInfo:=conStForeColorBox
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal, Me
End Sub

Private Sub mnuBlacknWhite_Click()
  ImageFilter conFltBlacknWhite
End Sub

Private Sub mnuBlur_Click()
  ImageFilter intFilter:=conFltBlur
End Sub

Private Sub mnuBrightness_Click()
  ImageFilter conFltBrightness
End Sub

Private Sub mnuBS_Click(Index As Integer)
  On Error GoTo ErrorHandler
  
  Dim i As Integer
  
  For i = 0 To mnuBS.Count - 1
    mnuBS(i).Checked = False
  Next
  intDrawStyle = Index
  mnuBS(Index).Checked = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuClear_Click()
  On Error GoTo ErrorHandler
  
  picPaint_DblClick
  picPaint.Picture = Nothing
  SetImageBuffer
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuCopy_Click()
  On Error GoTo ErrorHandler
  
  picClipboard.Picture = picSelect.Image
  mnuPaste.Enabled = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuCrop_Click()
  picSelect.Visible = False
  picPaint.Picture = picSelect.Image
  SetImageBuffer
  Form_Resize
  AdjustPaintResizeBox
End Sub

Private Sub mnuCut_Click()
  mnuDelete_Click
  mnuCopy_Click
End Sub

Private Sub mnuDarkness_Click()
  ImageFilter intFilter:=conFltDarkness
End Sub

Private Sub mnuDelete_Click()
  On Error GoTo ErrorHandler
  
  picSelect.Visible = False
  With picPaint
    'Remove the selection rectangle
    .DrawMode = vbXorPen
    .DrawStyle = vbDot
    .DrawWidth = 1
    picPaint.Line (picSelect.Left - 1, picSelect.Top - 1)- _
                  (picSelect.Left + picSelect.ScaleWidth, _
                   picSelect.Top + picSelect.ScaleHeight), _
                  vbBlack Xor picPaint.BackColor, B
    'Delete the selection area
    .DrawMode = vbCopyPen
    .DrawStyle = intDrawStyle
    If blnFirstMoving Then
      picPaint.Line (lngP1.x + varIIf(lngP1.x < lngP2.x, 1, -1), _
                     lngP1.y + varIIf(lngP1.y < lngP2.y, 1, -1))- _
                    (lngP2.x + varIIf(lngP2.x < lngP1.x, 1, -1), _
                     lngP2.y + varIIf(lngP2.y < lngP1.y, 1, -1)), _
                    picPaint.BackColor, BF
    End If
    .SetFocus
  End With
  picSelect.Visible = False
  mnuCut.Enabled = False
  mnuCopy.Enabled = False
  mnuDelete.Enabled = False
  mnuCrop.Enabled = False
  SetImageBuffer
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuDiffuse_Click()
  ImageFilter intFilter:=conFltDiffuse
End Sub

Private Sub mnuEdit_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuEffect_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuEmboss_Click()
  ImageFilter intFilter:=conFltEmboss
End Sub

Private Sub mnuExit_Click()
  On Error GoTo ErrorHandler

  Unload Me
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuFile_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuFillColor_Click()
  lblFillColor_DblClick
End Sub

Private Sub mnuFilter_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuFilterTools_Click(Index As Integer)
  On Error GoTo ErrorHandler
  
  Dim i As Integer
  
  For i = 0 To mnuFilterTools.Count - 1
    mnuFilterTools(i).Checked = False
  Next
  mnuFilterTools(Index).Checked = True
  intActiveFilterTool = Index
  picPaint.SetFocus
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuFlipHorizontal_Click()
  ImageEffect intEffect:=conEffFlipHorizontal
End Sub

Private Sub mnuFlipVertical_Click()
  ImageEffect intEffect:=conEffFlipVertical
End Sub

Private Sub mnuFont_Click()
  On Error GoTo ErrorHandler
  
  With cdlFonts
    'Set font dialog box properties with current paint area font properties
    .FontBold = picPaint.FontBold
    .FontItalic = picPaint.FontItalic
    .FontName = picPaint.FontName
    .FontSize = picPaint.FontSize
    .FontStrikethru = picPaint.FontStrikethru
    .FontUnderline = picPaint.FontUnderline
    .Color = picPaint.ForeColor
    'Open font dialog box
    .ShowFont
    'Set paint area and text box txtText font properties with properties in font
    '  dialog box
    picPaint.FontBold = .FontBold
    picPaint.FontItalic = .FontItalic
    picPaint.FontName = .FontName
    picPaint.FontSize = .FontSize
    picPaint.FontStrikethru = .FontStrikethru
    picPaint.FontUnderline = .FontUnderline
    picPaint.ForeColor = .Color
    txtText.FontBold = .FontBold
    txtText.FontItalic = .FontItalic
    txtText.FontName = .FontName
    txtText.FontSize = .FontSize
    txtText.FontStrikethru = .FontStrikethru
    txtText.FontUnderline = .FontUnderline
    txtText.ForeColor = .Color
    lblTextSize.FontBold = .FontBold
    lblTextSize.FontItalic = .FontItalic
    lblTextSize.FontName = .FontName
    lblTextSize.FontSize = .FontSize
    lblTextSize.FontStrikethru = .FontStrikethru
    lblTextSize.FontUnderline = .FontUnderline
    lblForeColor.BackColor = .Color
    txtText_KeyDown 0, 0
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuForegroundCOlor_Click()
  lblForeColor_DblClick
End Sub

Private Sub mnuFS_Click(Index As Integer)
  Dim i As Integer
  
  On Error GoTo ErrorHandler

  For i = 0 To mnuFS.Count - 1
    mnuFS(i).Checked = False
  Next
  intInsideFillStyle = Index
  mnuFS(Index).Checked = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuGrayBlacknWhite_Click()
  ImageFilter intFilter:=conFltGrayBlacknWhite
End Sub

Private Sub mnuGrayscale_Click()
  ImageFilter intFilter:=conFltGrayscale
End Sub

Private Sub mnuHelp_Click()
  UpdateStatusBar blnClear:=True
End Sub

Private Sub mnuInvertColors_Click()
  ImageEffect intEffect:=conEffInvertColors
End Sub

Private Sub mnuNew_Click()
  Dim i As Integer
  Dim intSave As Integer
  
  On Error GoTo ErrorHandler

  If blnPicChanged = True Then
    intSave = MsgBox("Do you want to save the changes?", _
                     vbYesNoCancel + vbExclamation)
  Else
    intSave = vbNo
  End If
  If intSave = vbYes Then
    mnuSave_Click
  End If
  If intSave <> vbCancel Then
    picZoom.Width = picPaint.Width
    picZoom.Height = picPaint.Height
    picZoom.Picture = Nothing
    ImageZoom blnNoZoom:=True
    picPaint.Picture = Nothing
    blnPicChanged = False
    strFileName = ""
    UpdateFormTitle
    blnDrawingPolygon = False
    ReDim lngPolygon(0)
    For i = 0 To 3
      imgBezier(i).Visible = False
    Next
    sngZoomFactor = 1
    AdjustPaintResizeBox
    ClearImageBuffer
    picSelect.Visible = False
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    mnuDelete.Enabled = False
    mnuCrop.Enabled = False
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuOpen_Click()
  Dim intSave As Integer
  
  On Error GoTo ErrorHandler
  
  If blnPicChanged Then
    intSave = MsgBox("Do you want to save the changes?", _
                     vbYesNoCancel + vbExclamation)
  Else
    intSave = vbNo
  End If
  If intSave = vbYes Then
    mnuSave_Click
  End If
  If intSave <> vbCancel Then
    cdlOpen.ShowOpen
    If cdlOpen.FileName <> "" Then
      blnPicChanged = False
      mnuNew_Click
      picPaint.Picture = LoadPicture(cdlOpen.FileName)
      strFileName = cdlOpen.FileName
      UpdateFormTitle
      ClearImageBuffer
      optTools_Click Index:=conTZoom
    End If
  End If
  Form_Resize
  AdjustPaintResizeBox
  Exit Sub

ErrorHandler:
  If Err.Number <> conErrCancel Then
    ShowErrMessage intErr:=conErrReadImage
  End If
End Sub

Private Sub mnuPaste_Click()
  On Error GoTo ErrorHandler
  
  picPaint_DblClick
  If Not blnFirstMoving Then
    PlaceSelection
  End If
  picSelect.Picture = picClipboard.Image
  picPaint.DrawStyle = vbDot
  blnFirstMoving = False
  If picSelect.Visible Then
    picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y), _
                  vbBlack Xor picPaint.BackColor, B
  End If
  picPaint.DrawMode = vbXorPen
  picPaint.DrawWidth = 1
  picSelect.Left = 0
  picSelect.Top = 0
  picPaint.Line (-1, -1)-(picClipboard.Width, picClipboard.Height), _
                vbBlack Xor picPaint.BackColor, B
  picSelect.Visible = True
  If intActiveTool <> conTSelect Then
    intActiveTool = conTSelect
    optTools(conTSelect).SetFocus
  End If
  mnuCut.Enabled = True
  mnuCopy.Enabled = True
  mnuDelete.Enabled = True
  mnuCrop.Enabled = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuPrint_Click()
  Dim strImgTmpFile As String

  On Error GoTo ErrorHandler
  
  cdlPrint.ShowPrinter
  Printer.Copies = cdlPrint.Copies
  strImgTmpFile = "temp.bmp"
  If blnFileExist(strImgTmpFile) Then
    Kill strImgTmpFile
  End If
  ImageZoom blnNoZoom:=True
  SavePicture picPaint.Image, strImgTmpFile
  picPaint.Picture = LoadPicture(strImgTmpFile)
  Kill strImgTmpFile
  Printer.PaintPicture picPaint, 0, 0
  Printer.EndDoc
  Exit Sub

ErrorHandler:
  If Err.Number <> conErrCancel Then
    ShowErrMessage intErr:=conErrPrint
  End If
End Sub

Private Sub mnuRedo_Click()
  On Error GoTo ErrorHandler
  
  ImageZoom blnNoZoom:=True
  'Remove selection
  If picSelect.Visible Then
    picSelect.Visible = False
    mnuCut.Enabled = False
    mnuCopy.Enabled = False
    mnuDelete.Enabled = False
    mnuCrop.Enabled = False
  End If
  'Set the current buffer index
  If intBufCur < conBufMax Then
    intBufCur = intBufCur + 1
  Else
    intBufCur = 0
  End If
  'Replace the paint area with image in picBuffer(intBufCur)
  picPaint.Picture = picBuffer(intBufCur).Image
  picPaint.Width = CLng(Left(picBuffer(intBufCur).Tag, _
                             Len(picBuffer(intBufCur).Tag) - 5))
  picPaint.Height = CLng(Right(picBuffer(intBufCur).Tag, 5))
  'Other settings
  mnuUndo.Enabled = True
  If intBufCur = intBufEnd Then
    mnuRedo.Enabled = False
  End If
  picPaint_DblClick
  AdjustPaintResizeBox
  Form_Resize
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuCrease_Click()
  ImageFilter intFilter:=conFltCrease
End Sub

Private Sub mnuReplaceColors_Click()
  ImageFilter intFilter:=conFltReplaceColors
End Sub

Private Sub mnuSnow_Click()
  ImageFilter intFilter:=conFltSnow
End Sub

Private Sub mnuResize125_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=1.25
End Sub

Private Sub mnuResize150_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=1.5
End Sub

Private Sub mnuResize175_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=1.75
End Sub

Private Sub mnuResize200_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=2
End Sub

Private Sub mnuResize25_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=0.25
End Sub

Private Sub mnuResize50_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=0.5
End Sub

Private Sub mnuResize75_Click()
  ImageEffect intEffect:=conEffResize, sngResizeFactor:=0.75
End Sub

Private Sub mnuResizeBoth_Click()
  On Error GoTo ErrorHandler

  mnuResizeBoth.Checked = True
  mnuResizeWidth.Checked = False
  mnuResizeHeight.Checked = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuResizeHeight_Click()
  On Error GoTo ErrorHandler

  mnuResizeBoth.Checked = False
  mnuResizeWidth.Checked = False
  mnuResizeHeight.Checked = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuResizeWidth_Click()
  On Error GoTo ErrorHandler

  mnuResizeBoth.Checked = False
  mnuResizeWidth.Checked = True
  mnuResizeHeight.Checked = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuRotate135_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=135
End Sub

Private Sub mnuRotate180_Click()
  ImageEffect intEffect:=conEffFlipHorizontal
  ImageEffect intEffect:=conEffFlipVertical
End Sub

Private Sub mnuRotate225_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=225
End Sub

Private Sub mnuRotate270_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=270
End Sub

Private Sub mnuRotate315_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=315
End Sub

Private Sub mnuRotate45_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=45
End Sub

Private Sub mnuRotate90_Click()
  ImageEffect intEffect:=conEffRotate, sngRotateAngle:=90
End Sub

Private Sub mnuRotateAntiClockwise_Click()
  On Error GoTo ErrorHandler

  mnuRotateClockwise.Checked = False
  mnuRotateAntiClockwise.Checked = True
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuRotateClockwise_Click()
  On Error GoTo ErrorHandler

  mnuRotateClockwise.Checked = True
  mnuRotateAntiClockwise.Checked = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuSave_Click()
  On Error GoTo ErrorHandler
  
  If strFileName = "" Then
    mnuSaveAs_Click
  Else
    ImageZoom blnNoZoom:=True
    SavePicture picPaint.Image, strFileName
    blnPicChanged = False
    UpdateFormTitle
  End If
  Exit Sub
  
ErrorHandler:
  ShowErrMessage intErr:=conErrWrite
End Sub

Private Sub mnuSaveAs_Click()
  On Error GoTo ErrorHandler
  
  cdlSave.ShowSave
  If cdlSave.FileName <> "" Then
    strFileName = cdlSave.FileName
    mnuSave_Click
  End If
  Exit Sub
  
ErrorHandler:
  If Err.Number = conErrPermission Then
    If ForceSave(strFileName) Then
      Resume
    End If
  ElseIf Err.Number <> conErrCancel Then
    ShowErrMessage intErr:=conErrWrite
  End If
End Sub

Private Sub mnuSharpen_Click()
  ImageFilter intFilter:=conFltSharpen
End Sub



Private Sub mnuUndo_Click()
  On Error GoTo ErrorHandler

  ImageZoom blnNoZoom:=True
  'Place the selection
  If picSelect.Visible Then
    PlaceSelection
    picPaint.SetFocus
  Else
    picPaint_DblClick
  End If
  'Set the current buffer index
  If intBufCur > 0 Then
    intBufCur = intBufCur - 1
  Else
    intBufCur = conBufMax
  End If
  'Replace the paint area with image in picBuffer(intBufCur)
  picPaint.Picture = picBuffer(intBufCur).Image
  picPaint.Width = CLng(Left(picBuffer(intBufCur).Tag, _
                             Len(picBuffer(intBufCur).Tag) - 5))
  picPaint.Height = CLng(Right(picBuffer(intBufCur).Tag, 5))
  'Other settings
  If intBufCur = intBufStart Then
    mnuUndo.Enabled = False
  End If
  mnuRedo.Enabled = True
  AdjustPaintResizeBox
  Form_Resize
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub mnuWave_Click()
  ImageFilter intFilter:=conFltWave
End Sub

Private Sub optTools_Click(Index As Integer)
  On Error GoTo ErrorHandler
  
  'Arrange draw width box and fill style box visibility
  Select Case intActiveTool
    Case conTAirBrush, conTArrow, conTCurve, conTEraser, _
         conTFilter, conTLine, conTPencil
      fraBrush.Visible = False
      fraOptDot.Visible = True
      fraOptFill.Visible = False
    Case conTRect, conTEllipse, conTRoundRect, conTPolygon
      fraBrush.Visible = False
      fraOptDot.Visible = True
      fraOptFill.Visible = True
    Case conTBrush
      fraBrush.Visible = True
      fraOptDot.Visible = True
      fraOptFill.Visible = False
    Case Else
      fraBrush.Visible = False
      fraOptDot.Visible = False
      fraOptFill.Visible = False
  End Select
  'Other settings
  If intActiveTool = conTFilter Then
    PopupMenu mnuTFilter
  End If
  If intActiveTool = conTZoom Then
    picZoom.Width = picPaint.Width
    picZoom.Height = picPaint.Height
    picZoom.Picture = picPaint.Image
  End If
  If intActiveTool <> conTSelect Then
    PlaceSelection
  End If
  If (intActiveTool <> conTPick) And (intActiveTool <> conTHand) Then
    ImageZoom blnNoZoom:=True
  End If
  UpdateStatusBar
  ChangePaintCursor
  picPaint.SetFocus
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub optTools_MouseDown(Index As Integer, Button As Integer, _
                               Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  If Button = vbLeftButton Then
    picPaint_DblClick
    intActiveTool = Index
    If intActiveTool = conTFilter Then
      PopupMenu mnuTFilter
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_DblClick()
  Dim i As Integer
  
  On Error GoTo ErrorHandler
  
  Select Case intActiveTool
    Case conTCurve
      If imgBezier(0).Visible Then
        DrawCurveBezier
        picPaint.DrawMode = vbCopyPen
        picPaint.ForeColor = lblForeColor.BackColor
        DrawCurveBezier blnComplete:=True
        For i = 0 To 3
          imgBezier(i).Visible = False
        Next
        SetImageBuffer
      End If
    Case conTPolygon
      If blnDrawingPolygon Then
        DrawPolygon blnComplete:=False
        DrawPolygon
        blnDrawingPolygon = False
        SetImageBuffer
      End If
    Case conTSelect
      PlaceSelection
  End Select
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim blnSuccess As Boolean

  On Error GoTo ErrorHandler

  If KeyCode = vbKeyReturn Then
    picPaint_DblClick
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_MouseDown(Button As Integer, _
                               Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  Dim i As Long
  
  If Button = vbLeftButton Then
    blnDrawing = True
    lngP1.x = x
    lngP1.y = y
    With picPaint
      If intActiveTool = conTSelect Then
        .DrawStyle = vbDot
        .DrawWidth = 1
      Else
        .DrawStyle = intDrawStyle
        .DrawWidth = intDot + 1
      End If
      Select Case intActiveTool
        Case conTAirBrush
          .DrawMode = vbCopyPen
          .ForeColor = lblForeColor.BackColor
          DrawAirBrush CInt(x), CInt(y), .DrawWidth * 4
        Case conTBrush
          .DrawMode = vbCopyPen
          .ForeColor = lblForeColor.BackColor
          .FillColor = lblForeColor.BackColor
          DrawBrush intBrushShape:=intBrushShape, x:=x, y:=y
        Case conTCurve
          If Not imgBezier(0).Visible Then
            .DrawMode = vbXorPen
            .ForeColor = picPaint.BackColor Xor lblForeColor.BackColor
            DrawCurveBezier blnCreate:=True, x:=x, y:=y
          End If
          lngP1.x = x
          lngP1.y = y
        Case conTEraser
          .DrawMode = vbCopyPen
          .ForeColor = .BackColor
          picPaint.Line (x, y)-(x + .DrawWidth, y - .DrawWidth), , B
        Case conTFill
          .DrawMode = vbCopyPen
          .FillColor = lblForeColor.BackColor
          .FillStyle = intInsideFillStyle
          mdlAPI.ExtFloodFill .hDC, x, y, .Point(x, y), 1
        Case conTFilter
          ImageFilter intFilter:=intActiveFilterTool, x:=CLng(x), y:=CLng(y)
        Case conTHand
          .ScaleMode = vbTwips
          .MouseIcon = LoadPicture(App.Path & "\Cursors\handgrab.cur")
          lngP1.x = (x * Screen.TwipsPerPixelX) + .Left
          lngP1.y = (y * Screen.TwipsPerPixelY) + .Top
          lngDragStart.x = .Left
          lngDragStart.y = .Top
          blnDrag = True
        Case conTPencil
          .DrawMode = vbCopyPen
          .ForeColor = lblForeColor.BackColor
          picPaint.Line (x, y)-(x, y), , B
        Case conTPick
          lblForeColor.BackColor = picPaint.Point(x, y)
        Case conTPolygon
          If Not blnDrawingPolygon Then
            blnDrawingPolygon = True
            ReDim lngPolygon(0)
            lngPolygon(0).x = x
            lngPolygon(0).y = y
          Else
            ReDim Preserve lngPolygon(UBound(lngPolygon) + 1)
            lngPolygon(UBound(lngPolygon)).x = x
            lngPolygon(UBound(lngPolygon)).y = y
            DrawPolygon blnComplete:=False
          End If
          .DrawMode = vbXorPen
          .FillStyle = vbFSTransparent
          .ForeColor = .BackColor Xor lblForeColor.BackColor
        Case conTText
          With txtText
            If Not .Visible Then
              .BackColor = picPaint.BackColor
              .ForeColor = lblForeColor.BackColor
              .Left = x
              .Top = y
              .Text = ""
              .Visible = True
              .SetFocus
            Else
              .Tag = "moving"
              .Move x, y
              .SetFocus
            End If
          End With
        Case Else
          If (intActiveTool = conTArrow) Or _
             (intActiveTool = conTSelect) Or (intActiveTool = conTLine) Then
            picPaint.Line (x, y)-(x, y)
          End If
          If intActiveTool = conTSelect Then
            .DrawWidth = 1
            PlaceSelection
          End If
          .DrawMode = vbXorPen
          If (intActiveTool = conTLine) Or _
             (intActiveTool = conTArrow) Or (intActiveTool = conTSelect) Then
            .ForeColor = .BackColor Xor lblForeColor.BackColor
            .FillStyle = vbFSTransparent
          Else
            Select Case intFillStyle
              Case conTsBorderOnly
                .FillStyle = vbFSTransparent
                .ForeColor = .BackColor Xor lblForeColor.BackColor
              Case conTsBorderFill
                .FillStyle = intInsideFillStyle
                .ForeColor = .BackColor Xor lblForeColor.BackColor
                .FillColor = .BackColor Xor lblFillColor.BackColor
              Case conTsFillOnly
                .FillStyle = intInsideFillStyle
                .ForeColor = .BackColor Xor lblFillColor.BackColor
                .FillColor = .BackColor Xor lblFillColor.BackColor
            End Select
          End If
          lngP2 = lngP1
      End Select
    End With
  ElseIf (Button = vbRightButton) Then
    If intActiveTool = conTPick Then
      lblFillColor.BackColor = picPaint.Point(x, y)
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrDrawing
End Sub

Private Sub picPaint_MouseMove(Button As Integer, _
                               Shift As Integer, x As Single, y As Single)
  Dim intHscPaintValue As Integer       'adjusted horizontal and vertical scroll
  Dim intVscPaintValue As Integer       '                bar value for hand tool
  
  On Error GoTo ErrorHandler
  
  If Button = vbLeftButton Then
    If blnDrawing Then
      With picPaint
        Select Case intActiveTool
          Case conTAirBrush
            DrawAirBrush CInt(x), CInt(y), .DrawWidth * 4
          Case conTArrow
            DrawArrow lngP1.x, lngP1.y, lngP2.x, lngP2.y
            AdjustP2 x:=x, y:=y, Shift:=Shift, blnEnableCtrl:=True
            DrawArrow lngP1.x, lngP1.y, lngP2.x, lngP2.y
          Case conTBrush
            .DrawMode = vbCopyPen
            .ForeColor = lblForeColor.BackColor
            .FillColor = lblForeColor.BackColor
            DrawBrush intBrushShape:=intBrushShape, x:=x, y:=y
          Case conTCurve
            DrawCurveBezier
            imgBezier(0).Top = imgBezier(0).Top + (y - lngP1.y)
            imgBezier(0).Left = imgBezier(0).Left + (x - lngP1.x)
            imgBezier(1).Top = imgBezier(1).Top + (y - lngP1.y)
            imgBezier(1).Left = imgBezier(1).Left + (x - lngP1.x)
            imgBezier(2).Top = imgBezier(2).Top + (y - lngP1.y)
            imgBezier(2).Left = imgBezier(2).Left + (x - lngP1.x)
            imgBezier(3).Top = imgBezier(3).Top + (y - lngP1.y)
            imgBezier(3).Left = imgBezier(3).Left + (x - lngP1.x)
            DrawCurveBezier
            lngP1.x = x
            lngP1.y = y
          Case conTEllipse
            If (lngP2.x <> lngP1.x) Then
              picPaint.Circle ((lngP1.x + lngP2.x) / 2, _
                                 (lngP1.y + lngP2.y) / 2), _
                               varIIf(Abs(lngP2.x - lngP1.x) > _
                                        Abs(lngP2.y - lngP1.y), _
                                      Abs(lngP2.x - lngP1.x) / 2, _
                                      Abs(lngP2.y - lngP1.y) / 2), , , , _
                               Abs((lngP2.y - lngP1.y) / _
                                   (lngP2.x - lngP1.x))
            End If
            AdjustP2 x:=x, y:=y, Shift:=Shift
            If (lngP2.x <> lngP1.x) Then
              picPaint.Circle ((lngP1.x + lngP2.x) / 2, _
                                 (lngP1.y + lngP2.y) / 2), _
                               varIIf(Abs(lngP2.x - lngP1.x) > _
                                        Abs(lngP2.y - lngP1.y), _
                                      Abs(lngP2.x - lngP1.x) / 2, _
                                      Abs(lngP2.y - lngP1.y) / 2), , , , _
                               Abs((lngP2.y - lngP1.y) / _
                                   (lngP2.x - lngP1.x))
            End If
          Case conTEraser
            picPaint.Line (x, y)-(x + .DrawWidth, y - .DrawWidth), , B
          Case conTFilter
            ImageFilter intFilter:=intActiveFilterTool, x:=CLng(x), y:=CLng(y)
          Case conTHand
            If blnDrag Then
              If hscPaint.Visible Then
                intHscPaintValue = lngDragStart.x - _
                                   (lngP1.x - (x + picPaint.Left))
                intHscPaintValue = hscPaint.Value + _
                                   ((picPaint.Left - intHscPaintValue) / _
                                    Screen.TwipsPerPixelX)
                If intHscPaintValue < hscPaint.Min Then
                  hscPaint.Value = hscPaint.Min
                ElseIf intHscPaintValue > hscPaint.Max Then
                  hscPaint.Value = hscPaint.Max
                Else
                  hscPaint.Value = intHscPaintValue
                End If
              End If
              If vscPaint.Visible Then
                intVscPaintValue = lngDragStart.y - _
                                   (lngP1.y - (y + picPaint.Top))
                intVscPaintValue = vscPaint.Value + _
                                   ((picPaint.Top - intVscPaintValue) / _
                                    Screen.TwipsPerPixelY)
                If intVscPaintValue < vscPaint.Min Then
                  vscPaint.Value = vscPaint.Min
                ElseIf intVscPaintValue > vscPaint.Max Then
                  vscPaint.Value = vscPaint.Max
                Else
                  vscPaint.Value = intVscPaintValue
                End If
              End If
              picPaint.Refresh
            End If
          Case conTLine
            picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y)
            AdjustP2 x:=x, y:=y, Shift:=Shift, blnEnableCtrl:=True
            picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y)
          Case conTPencil
            lngP2 = lngP1
            lngP1.x = x
            lngP1.y = y
            picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y)
          Case conTPolygon
            If UBound(lngPolygon) = 0 Then
              ReDim Preserve lngPolygon(UBound(lngPolygon) + 1)
            Else
              DrawPolygon blnComplete:=False
            End If
            lngPolygon(UBound(lngPolygon)).x = x
            lngPolygon(UBound(lngPolygon)).y = y
            DrawPolygon blnComplete:=False
          Case conTRect
            If (lngP1.x <> lngP2.x) Or (lngP1.y <> lngP2.y) Then
              picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y), , B
            End If
            AdjustP2 x:=x, y:=y, Shift:=Shift
            picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y), , B
          Case conTRoundRect
            mdlAPI.RoundRect picPaint.hDC, _
                             lngP1.x, lngP1.y, lngP2.x, lngP2.y, 10, 10
            AdjustP2 x:=x, y:=y, Shift:=Shift
            mdlAPI.RoundRect picPaint.hDC, _
                             lngP1.x, lngP1.y, lngP2.x, lngP2.y, 10, 10
            .Refresh
          Case conTSelect
            picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y), _
                          vbBlack Xor picPaint.BackColor, B
            AdjustP2 x:=x, y:=y, Shift:=Shift
            picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y), _
                          vbBlack Xor picPaint.BackColor, B
        End Select
      End With
    End If
  End If
  UpdateStatusBar x:=x, y:=y
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_MouseUp(Button As Integer, _
                             Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  If Button = vbLeftButton Then
    If blnDrawing Then
      lngP2.x = x
      lngP2.y = y
      Select Case intActiveTool
        Case conTArrow, conTEllipse, conTLine, conTRect, conTRoundRect
          With picPaint
            .DrawMode = vbCopyPen
            If intActiveTool = conTLine Then
              .ForeColor = lblForeColor.BackColor
            Else
              .ForeColor = .BackColor Xor .ForeColor
              .FillColor = .BackColor Xor .FillColor
            End If
          End With
          Select Case intActiveTool
            Case conTArrow
              AdjustP2 x:=x, y:=y, Shift:=Shift, blnEnableCtrl:=True
              DrawArrow lngP1.x, lngP1.y, lngP2.x, lngP2.y
            Case conTEllipse
              AdjustP2 x:=x, y:=y, Shift:=Shift
              If (lngP2.x <> lngP1.x) Then
                picPaint.Circle ((lngP1.x + lngP2.x) / 2, _
                                   (lngP1.y + lngP2.y) / 2), _
                                 varIIf(Abs(lngP2.x - lngP1.x) > _
                                          Abs(lngP2.y - lngP1.y), _
                                        Abs(lngP2.x - lngP1.x) / 2, _
                                        Abs(lngP2.y - lngP1.y) / 2), , , , _
                                 Abs((lngP2.y - lngP1.y) / _
                                     (lngP2.x - lngP1.x))
              End If
            Case conTLine
              AdjustP2 x:=x, y:=y, Shift:=Shift, blnEnableCtrl:=True
              picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y)
            Case conTRect
              AdjustP2 x:=x, y:=y, Shift:=Shift
              If (lngP1.x <> lngP2.x) Or (lngP1.y <> lngP2.y) Then
                picPaint.Line (lngP1.x, lngP1.y)- _
                              (lngP2.x, lngP2.y), , B
              End If
            Case conTRoundRect
              AdjustP2 x:=x, y:=y, Shift:=Shift
              mdlAPI.RoundRect picPaint.hDC, _
                               lngP1.x, lngP1.y, lngP2.x, lngP2.y, 10, 10
          End Select
        Case conTHand
          blnDrag = False
          picPaint.ScaleMode = vbPixels
          picPaint.MouseIcon = LoadPicture(App.Path & "\Cursors\handflat.cur")
        Case conTSelect
          With picSelect
            If (Abs(lngP2.x - lngP1.x) > 1) And _
               (Abs(lngP2.y - lngP1.y) > 1) Then
              AdjustP2 x:=x, y:=y, Shift:=Shift
              .Width = Abs(lngP2.x - lngP1.x) - 1
              .Height = Abs(lngP2.y - lngP1.y) - 1
              .Left = IIf(lngP1.x <= lngP2.x, lngP1.x, lngP2.x) + 1
              .Top = IIf(lngP1.y <= lngP2.y, lngP1.y, lngP2.y) + 1
              .Visible = True
              .Picture = Nothing
              .PaintPicture picPaint.Image, 0, 0, _
                            .Width, .Height, .Left, .Top, .Width, .Height
              mnuCut.Enabled = True
              mnuCopy.Enabled = True
              mnuDelete.Enabled = True
              mnuCrop.Enabled = True
              blnFirstMoving = True
            Else
              .Visible = False
              picPaint.Line (lngP1.x, lngP1.y)-(lngP2.x, lngP2.y), _
                            vbBlack Xor picPaint.BackColor, B
              mnuCut.Enabled = False
              mnuCopy.Enabled = False
              mnuDelete.Enabled = False
              mnuCrop.Enabled = False
              blnFirstMoving = False
            End If
          End With
          picPaint.DrawWidth = intDot + 1
        Case conTZoom
          If sngZoomFactor = 1 Then
            picZoom.Width = picPaint.Width
            picZoom.Height = picPaint.Height
            picZoom.Picture = picPaint.Image
          End If
          If Shift <> vbCtrlMask Then
            'Zoom in
            If ((picZoom.Width * sngZoomFactor * conZoomFactor * 2) <= _
                (mdlEffect.conMaxImageWidth * 2)) And _
               ((picZoom.Height * sngZoomFactor * conZoomFactor * 2) <= _
                (mdlEffect.conMaxImageHeight * 2)) Then
              sngZoomFactor = sngZoomFactor * conZoomFactor
              ImageZoom x:=CLng(x * Screen.TwipsPerPixelX * conZoomFactor), _
                        y:=CLng(y * Screen.TwipsPerPixelY * conZoomFactor)
            End If
          Else
            'Zoom out
            sngZoomFactor = sngZoomFactor / conZoomFactor
            ImageZoom x:=CLng(x * Screen.TwipsPerPixelX / conZoomFactor), _
                      y:=CLng(y * Screen.TwipsPerPixelY / conZoomFactor)
          End If
      End Select
      blnDrawing = False
      If (intActiveTool <> conTText) And (intActiveTool <> conTSelect) And _
         (intActiveTool <> conTPolygon) And (intActiveTool <> conTCurve) And _
         (intActiveTool <> conTZoom) Then
        SetImageBuffer
      End If
    End If
  End If
  UpdateStatusBar
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaint_Resize()
  blnResize = True
End Sub

Private Sub picPaintResize_MouseDown(Index As Integer, Button As Integer, _
                                     Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  'Start the drag operation on picPaintResize(Index)
  lngDragStart.x = CLng(x)
  lngDragStart.y = CLng(y)
  blnDrag = True
  blnResize = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaintResize_MouseMove(Index As Integer, Button As Integer, _
                                     Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  'Move picPaintResize(Index) for the drag operation and resize the paint area
  '  to match picPaintResize(Index) position
  If blnDrag Then
    With picPaintResize(Index)
      If Index <> conResizeNS Then
        If (picPaint.Width + (x - lngDragStart.x)) > 0 Then
          .Left = .Left + (x - lngDragStart.x)
          picPaint.Width = picPaint.Width + (x - lngDragStart.x)
        End If
      End If
      If Index <> conResizeWE Then
        If (picPaint.Height + (y - lngDragStart.y)) > 0 Then
          .Top = .Top + (y - lngDragStart.y)
          picPaint.Height = picPaint.Height + (y - lngDragStart.y)
        End If
      End If
    End With
    AdjustPaintResizeBox
    Form_Resize
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picPaintResize_MouseUp(Index As Integer, Button As Integer, _
                                   Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  'End the drag operation on picPaintResize(Index)
  blnDrag = False
  If blnResize Then
    SetImageBuffer
  End If
  blnResize = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picSelect_MouseDown(Button As Integer, _
                                Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  If Button = vbLeftButton Then
    'Start the drag operation on the selection object
    blnMoving = True
    With picSelect
      picPaint.DrawWidth = 1
      If blnFirstMoving And (Shift <> vbCtrlMask) Then
        'Erase the drawing behind the selection object
        picPaint.DrawStyle = intDrawStyle
        picPaint.DrawMode = vbCopyPen
        picPaint.Line (.Left, .Top)-(.Left + .Width - 1, .Top + .Height - 1), _
                      picPaint.BackColor, BF
        blnFirstMoving = False
      End If
      picPaint.DrawStyle = vbDot
      picPaint.DrawMode = vbXorPen
      picPaint.Line (.Left - 1, .Top - 1)- _
                    (.Left + .Width, .Top + .Height), _
                    vbBlack Xor picPaint.BackColor, B
      lngP1.x = x
      lngP1.y = y
    End With
  End If
  UpdateStatusBar
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picSelect_MouseMove(Button As Integer, _
                                Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  'Move the selection object for the drag operation
  If (Button = vbLeftButton) And blnMoving Then
    lngP2.x = x
    lngP2.y = y
    picSelect.Left = picSelect.Left + (lngP2.x - lngP1.x)
    picSelect.Top = picSelect.Top + (lngP2.y - lngP1.y)
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub picSelect_MouseUp(Button As Integer, _
                              Shift As Integer, x As Single, y As Single)
  On Error GoTo ErrorHandler
  
  'End the drag operation on picSelect
  If Button = vbLeftButton Then
    With picSelect
      picPaint.Line (.Left - 1, .Top - 1)- _
                    (.Left + .Width, .Top + .Height), _
                    vbBlack Xor picPaint.BackColor, B
    End With
    blnFirstMoving = False
    blnMoving = False
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Place selection image in picSelect to the paint area
' Assumptions: * These components exists in this form:
'                - picPaint
'                - picSelect
'              * Meet assumptions in this procedure:
'                  SetImageBuffer
' Effects    : * picSelect.Visible = False
'              * The selection rectangle has been erased
'              * Effects from SetImageBUffer
'              * Menu "Delete" is not enabled
' Inputs     : -
' Returns    : -
Private Sub PlaceSelection()
  On Error GoTo ErrorHandler

  With picSelect
    If .Visible Then
      .Visible = False
      picPaint.PaintPicture .Image, .Left, .Top
      'Erase the selection rectangle
      picPaint.DrawMode = vbXorPen
      picPaint.DrawWidth = 1
      picPaint.Line (.Left - 1, .Top - 1)-(.Left + .Width, .Top + .Height), _
                    vbBlack Xor picPaint.BackColor, B
      If Not blnFirstMoving Then
        SetImageBuffer
      End If
      mnuCopy.Enabled = False
      mnuCut.Enabled = False
      mnuCrop.Enabled = False
      mnuDelete.Enabled = False
    End If
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Safe the image on paint area (picPaint) to image buffer array
'              (picBuffer)
' Assumptions: These components exists in this form:
'                - picPaint
'                - picBuffer
' Effects    : * Image in paint area has been saved to image buffer array
'              * Buffer pointer (intBufStart) has been set to the next buffer
' Inputs     : -
' Returns    : -
Public Sub SetImageBuffer()
  On Error GoTo ErrorHandler

  If intBufCur < conBufMax Then
    intBufCur = intBufCur + 1
  Else
    intBufCur = 0
  End If
  If intBufCur > picBuffer.UBound Then
    Load picBuffer(intBufCur)
  End If
  picBuffer(intBufCur).Picture = picPaint.Image
  picBuffer(intBufCur).Tag = CStr((picPaint.Width * 100000) + picPaint.Height)
  intBufEnd = intBufCur
  If intBufStart = intBufEnd Then
    If intBufStart < conBufMax Then
      intBufStart = intBufStart + 1
    Else
      intBufStart = 0
    End If
  End If
  blnPicChanged = True
  mnuUndo.Enabled = True
  mnuRedo.Enabled = False
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub txtText_DblClick()
  On Error GoTo ErrorHandler
  
  With txtText
    picPaint.CurrentX = .Left
    picPaint.CurrentY = .Top
    picPaint.ForeColor = lblForeColor.BackColor
    picPaint.Print .Text
    .Visible = False
    SetImageBuffer
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorHandler
  
  With txtText
    lblTextSize.Caption = .Text & "M"
    .Width = lblTextSize.Width
    .Height = lblTextSize.Height
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorHandler
  
  With txtText
    Select Case KeyCode
      Case vbKeyReturn
        txtText_DblClick
      Case vbKeyEscape
        .Visible = False
      Case Else
        lblTextSize.Caption = .Text & "O"
        .Width = lblTextSize.Width
        .Height = lblTextSize.Height
    End Select
    If Not .Visible Then
      picPaint.SetFocus
    End If
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub txtText_LostFocus()
  On Error GoTo ErrorHandler
  
  With txtText
    If (.Visible) And (.Tag <> "moving") Then
      txtText_KeyUp vbKeyReturn, 0
    End If
    .Tag = ""
  End With
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Update certain drawing to match drawing properties changes
'              (dot width, foreground color, etc.)
' Assumption : This global variable has been initiated:
'                intActiveTool
' Effect     : As specified
' Inputs     : -
' Returns    : -
Private Sub UpdateDrawing()
  On Error GoTo ErrorHandler
  
  Select Case intActiveTool
    Case conTCurve
      DrawCurveBezier
    Case conTPolygon
      If blnDrawingPolygon Then
        DrawPolygon blnComplete:=False, blnOnlyDrawLastLine:=False
      End If
  End Select
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Update form's title based on file name strFileName
' Assumptions: This global variable has been initiated:
'                strFileName
'              This global constant has been initiated:
'                conProgramTitle (the title of this program)
' Effects    : The form's title has been updated
' Inputs     : -
' Returns    : -
Private Sub UpdateFormTitle()
  On Error GoTo ErrorHandler
  
  If strFileName <> "" Then
    Me.Caption = strFileName & " - " & conProgramTitle
  Else
    Me.Caption = "untitled - " & conProgramTitle
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

' Purpose    : Update the status bar sta content
' Assumptions: These components exists in this form:
'                sta, imgBezier()
'              This global variable has been initiated:
'                intActiveTool, blnFirstMoving, blnDrawing, blnDrawingPolygon
' Effect     : The status bar content has been updated
' Inputs     : * intArea (the new content of the status bar)
'              * X, Y (current drawing coordinates)
'              * blnClear (condition to remove all texts in status bar)
' Return     : -
Public Sub UpdateStatusBar(Optional intInfo As enmStatusBar = conStPaintArea, _
                           Optional x As Single, Optional y As Single, _
                           Optional intPercentage As Integer, _
                           Optional blnClear As Boolean = False)
  On Error GoTo ErrorHandler
  
  If blnClear Then
    sta.Panels(1).Text = ""
    sta.Panels(2).Text = ""
    sta.Panels(3).Text = ""
  Else
    'First panel
    With sta.Panels(1)
      Select Case intInfo
        Case conStPaintArea
          Select Case intActiveTool
            Case conTAirBrush
              .Text = "Draws using an airbrush with the selected airbrush size"
            Case conTArrow
              If Not blnDrawing Then
                .Text = "Draws an arrow with the selected arrow width"
              Else
                .Text = "Press and hold down " & _
                        "CTRL to draw a horizontal or vertical arrow; " & _
                        "SHIFT to draw a 45-degree arrow"
              End If
            Case conTBrush
              .Text = "Draws using a brush with the selected shape"
            Case conTCurve
              If Not imgBezier(0).Visible Then
                .Text = "Draws a bezier curve with the selected curve width"
              Else
                .Text = "Press ENTER or double-click " & _
                        "to finish drawing the curve"
              End If
            Case conTEllipse
              If Not blnDrawing Then
                .Text = "Draws an ellips " & _
                        "with the selected outline width and fill style"
              Else
                .Text = "Press and hold down SHIFT to draw a circle"
              End If
            Case conTEraser
              .Text = "Erases a partion of the picture " & _
                      "using the selected eraser width"
            Case conTFilter
              .Text = "Apply the selected filter to the image"
            Case conTFill
              .Text = "Fills an area"
            Case conTHand
              .Text = "Pan to see other part of the picture"
            Case conTLine
              If Not blnDrawing Then
                .Text = "Draws a straight line with the selected line width"
              Else
                .Text = "Press and hold down " & _
                        "CTRL to draw a horizontal or vertical line; " & _
                        "SHIFT to draw a 45-degree line"
              End If
            Case conTPencil
              .Text = "Draws using a pencil with the selected dot size"
            Case conTPick
              .Text = "Picks up a foreground color (click) or " & _
                      "background color (right-click) " & _
                      "from the picture for drawing"
            Case conTPolygon
              If Not blnDrawingPolygon Then
                .Text = "Draws a polygon " & _
                        "with the selected outline width and fill area"
              Else
                .Text = "Press ENTER or double-click " & _
                        "to finish drawing the polygon"
              End If
            Case conTRect
              If Not blnDrawing Then
                .Text = "Draws a rectangle " & _
                        "with the selected outline width and fill style"
              Else
                .Text = "Press and hold down SHIFT to draw a square"
              End If
            Case conTRoundRect
              If Not blnDrawing Then
                .Text = "Draws a rounded rectangle " & _
                        "with the selected outline width and fill style"
              Else
                .Text = "Press and hold down SHIFT to draw a rounded-square"
              End If
            Case conTSelect
              If blnFirstMoving Then
                .Text = "Press and hold down CTRL " & _
                        "before moving the selection to copy it"
              ElseIf Not blnDrawing Then
                .Text = "Selects a rectangular part of the picture " & _
                        "to move or delete"
              Else
                .Text = "Press and hold down SHIFT to select a square part"
              End If
            Case conTText
              If Not txtText.Visible Then
                .Text = "Insert text into the picture"
              Else
                .Text = "Press ENTER or double-click " & _
                        "to finish inserting the text"
              End If
            Case conTZoom
              .Text = "Zoom in or zoom out the image 1.25x " & _
                      "(press and hold down CTRL to zoom out)"
          End Select
        Case conStColorBox
          .Text = "Click to set the foreground color; " & _
                               "Right-click to set the background color"
        Case conStForeColorBox
          .Text = "Double-click " & _
                  "to set the foreground color with custom color"
        Case conStBackColorBox
          .Text = "Double-click " & _
                  "to set the background color with custom color"
        Case conStFiltering
          .Text = "Filtering... " & _
                 "(" & CStr(intPercentage) & "% complete)"
        Case conStRetrieveingColor
          .Text = "Retrieving color information... " & _
                  "(" & CStr(intPercentage) & "% complete)"
        Case Else
          .Text = ""
      End Select
    End With
    'Second and third panels
    If intInfo = conStPaintArea Then
      If blnDrawing And _
         ((intActiveTool = conTArrow) Or (intActiveTool = conTEllipse) Or _
          (intActiveTool = conTLine) Or (intActiveTool = conTRect) Or _
          (intActiveTool = conTRoundRect) Or (intActiveTool = conTSelect)) Then
        sta.Panels(3).Text = CStr(lngP2.x - lngP1.x) & "x" & _
                             CStr(lngP2.y - lngP1.y)
      Else
        sta.Panels(2).Text = CStr(x) & "," & CStr(y)
        sta.Panels(3).Text = ""
      End If
    Else
      sta.Panels(2).Text = ""
      sta.Panels(3).Text = ""
    End If
  End If
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub

Private Sub vscPaint_Change()
  Dim lngPicPaintTop As Long
  
  On Error GoTo ErrorHandler
  
  lngPicPaintTop = -(CLng(vscPaint.Value) * 10)
  picPaint.Top = lngPicPaintTop
  AdjustPaintResizeBox
  Exit Sub

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Sub



Private Sub cboType_Click()
    
    sstType.Tab = cboType.ListIndex
    
    Select Case cboType.ListIndex
        Case 0
            txtTitle.Text = "S_TEXT" & gblCtrlIdx
        Case 1
            txtTitle.Text = "D_TEXT" & gblCtrlIdx
        Case 2
            txtTitle.Text = "S_Image" & gblCtrlIdx
        Case 3
            txtTitle.Text = "D_Image" & gblCtrlIdx
        Case 4
            txtTitle.Text = "BARCODE" & gblCtrlIdx
        Case 5
            txtTitle.Text = "LINE" & gblCtrlIdx
            txtLineHSize.Text = "1"
    End Select
    
    txtXpos.Text = 1
    txtYpos.Text = 10
    
End Sub




'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 동적 생성 컨트롤에서의 이벤트 처리
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub ClsEventMonitor_EventRaised(EventObject As ClassEventObject, ByVal StrEventName As String)

    Dim StrEvent        As String
    Dim obj             As Object
    Dim val1            As Variant
    
    On Error Resume Next

    ' 실제 이벤트가 발생한 Object
    Set obj = EventObject.EventObject

    StrEvent = ""
    StrEvent = StrEvent & Format(Now, "HH:MM:SS") & " "
    StrEvent = StrEvent & obj.Name & " - " & StrEventName & "("
    
    ' 파라미터 정보
    For Each val1 In EventObject.Params
        StrEvent = StrEvent & CStr(val1) & ", "
    Next

    If Right(StrEvent, 2) = ", " Then
        StrEvent = Left(StrEvent, Len(StrEvent) - 2)
    End If

    StrEvent = StrEvent & "" & ")"
    
    ' 이벤트 로그
    List1.AddItem StrEvent, 0

End Sub

Private Sub cmdDelobj_Click()
    Dim intRow          As Integer
    Dim strObjType      As Variant
    Dim strObjName      As Variant
    Dim strObjRotate    As Variant
    
    Me.Controls(txtTag.Text).Visible = False
    
    With spdList
        For intRow = 1 To .MaxRows
            .Row = intRow
            Call .GetText(2, intRow, strObjType)
            Call .GetText(28, intRow, strObjName)
            '
            If strObjType = sstType.Tab And strObjName = Trim(txtTag.Text) Then
                .Action = ActionDeleteRow
                .MaxRows = .MaxRows - 1
                Exit For
            End If
        Next
    End With

End Sub

Private Sub cmdDevide_Click(Index As Integer)
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim strBuf()    As String
    
    intMode = 2
    
    If Index = 0 Then
        If txtDevide.Text = "0.1" Then
            txtDevide.Text = "0.1"
        Else
            txtDevide.Text = txtDevide.Text - 0.1
        End If
    Else
        txtDevide.Text = txtDevide.Text + 0.1
    End If
    gDevide = txtDevide.Text
    
    ' 컬렉션 초기화
    Set m_ColCommandButton = Nothing
    Set m_ColCommandButton = New Collection
    
    With spdList
        sstType.Visible = False
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = 1
            Erase strBuf
            If Trim(.Text) <> "" Then
                ReDim Preserve strBuf(.MaxCols) As String
                For intCol = 1 To .MaxCols
                    .Col = intCol
                    strBuf(intCol - 1) = Trim(.Text)
                Next
                Call MakeLayout(strBuf)
                Erase strBuf
            End If
        Next
        sstType.Visible = True
    End With

End Sub

'-- 폰트 설정
Private Sub cmdFont_Click(Index As Integer)
 
    'Cancel을 True로 설정합니다.
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    
    'Flags 속성을 설정합니다.
    CommonDialog1.flags = cdlCFEffects Or cdlCFBoth
    
    '폰트 속성을 설정합니다.[Default]
    CommonDialog1.FontName = "굴림"
    CommonDialog1.FontSize = 9
    
    '[글꼴] 대화 상자를 표시합니다.
    CommonDialog1.ShowFont
    txtFontName(Index).Text = CommonDialog1.FontName
    txtFontSize(Index).Text = CommonDialog1.FontSize
    chkFontBold(Index).Value = IIf(CommonDialog1.FontBold = True, 1, 0)
    chkFontItalic(Index).Value = IIf(CommonDialog1.FontItalic = True, 1, 0)
    chkFontUnder(Index).Value = IIf(CommonDialog1.FontUnderline = True, 1, 0)

    Exit Sub

ErrHandler:
  '" 사용자가 [취소] 단추를 눌렀습니다.
  Exit Sub
  
End Sub

'-- 이미지 경로 설정
Private Sub cmdImage_Click(Index As Integer)

    Dim sFile As String
    sFile = ShowOpen("JPG파일(*.jpg)|*.jpg", App.Path & "\" & gImage)
    If sFile <> "" Then
        txtImageName(Index).Text = sFile
        If Index = 0 Then
            Didim_SImg.Picture = LoadPicture(txtImageName(Index).Text)
            txtImageWSize(Index).Text = Round(Didim_SImg.Width / gScaleCal, 0)
            txtImageHSize(Index).Text = Round(Didim_SImg.Height / gScaleCal, 0)
            
            txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
            txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text
            
            txtImageDevide(Index).SetFocus
        Else
            Didim_DImg.Picture = LoadPicture(txtImageName(Index).Text)
            txtImageWSize(Index).Text = Round(Didim_DImg.Width / gScaleCal, 0)
            txtImageHSize(Index).Text = Round(Didim_DImg.Height / gScaleCal, 0)
        
            txtImageWSize(Index + 2).Text = txtImageWSize(Index).Text
            txtImageHSize(Index + 2).Text = txtImageHSize(Index).Text
            
            txtImageDevide(Index).SetFocus
        End If
    Else
'        MsgBox "You pressed cancel"
    End If




'
'
'Dim x
'    'Cancel을 True로 설정합니다.
'    CommonDialog1.CancelError = True
'    On Error GoTo ErrHandler
'
'    'Flags 속성을 설정합니다.
'    CommonDialog1.flags = cdlCFEffects Or cdlCFBoth
'
'    '경로 속성을 설정합니다.
'    CommonDialog1.InitDir = App.Path & "\" & gImage
'
'    CommonDialog1.Filter = "JPG파일(*.jpg)|*.jpg"
'
'    '[파일] 대화 상자를 표시합니다.
'    CommonDialog1.ShowOpen
'    txtImageName(Index).Text = CommonDialog1.FileName
'
'    If Index = 0 Then
'        Didim_SImg.Picture = LoadPicture(txtImageName(Index).Text)
'        txtImageWSize(Index).Text = Round(Didim_SImg.Width / gScaleCal, 0)
'        txtImageHSize(Index).Text = Round(Didim_SImg.Height / gScaleCal, 0)
'    Else
'        Didim_DImg.Picture = LoadPicture(txtImageName(Index).Text)
'        txtImageWSize(Index).Text = Round(Didim_DImg.Width / gScaleCal, 0)
'        txtImageHSize(Index).Text = Round(Didim_DImg.Height / gScaleCal, 0)
'    End If
'
'    Exit Sub
'
'ErrHandler:
'  '" 사용자가 [취소] 단추를 눌렀습니다.
'  Exit Sub

End Sub

Private Sub MakeSpdSaveList(obj As Object, idx As Integer)
    
    With spdList
        .MaxRows = .MaxRows + 1
        .Action = ActionActiveCell
        Select Case idx
        Case 0, 1
            .SetText 1, .MaxRows, .MaxRows - 1                                      '설정순번
            .SetText 2, .MaxRows, idx                                               '항목구분
            .SetText 3, .MaxRows, txtTitle.Text                                     '항목명
            .SetText 4, .MaxRows, txtXpos.Text                                      'X1좌표
            .SetText 5, .MaxRows, 0                                                 'X2좌표
            .SetText 6, .MaxRows, txtYpos.Text                                      'Y1좌표
            .SetText 7, .MaxRows, 0                                                 'Y2좌표
            .SetText 8, .MaxRows, txtFontName(idx).Text                             '폰트명
            .SetText 9, .MaxRows, txtFontSize(idx).Text                             '폰트크기
            .SetText 10, .MaxRows, IIf(chkFontBold(idx).Value = "0", "0", "1")      '폰트굵게
            .SetText 11, .MaxRows, IIf(chkFontUnder(idx).Value = "0", "0", "1")     '폰트밑줄
            .SetText 12, .MaxRows, IIf(chkFontItalic(idx).Value = "0", "0", "1")    '폰트기울게
            .SetText 13, .MaxRows, "0"                                              '폰트회전
            .SetText 14, .MaxRows, "0"                                              '바코드종류
            .SetText 15, .MaxRows, "0"                                              '바코드폭
            .SetText 16, .MaxRows, "0"                                              '바코드회전
            .SetText 17, .MaxRows, ""                                               '이미지경로
            .SetText 18, .MaxRows, "0"                                              '라인회전
            .SetText 19, .MaxRows, "0"                                              '라인두께
            .SetText 20, .MaxRows, "0"                                              '라인폭
            .SetText 21, .MaxRows, IIf(chkPrint.Value = "1", "0", "1")              '출력여부
            .SetText 22, .MaxRows, txtContent(idx).Text                             '출력값
            .SetText 23, .MaxRows, gScaleCal                                              'X좌표 보정값
            .SetText 24, .MaxRows, gScaleCal                                              'Y좌표 보정값
            .SetText 25, .MaxRows, txtPaperHSize.Text                               '용지높이
            .SetText 26, .MaxRows, txtPaperWSize.Text                               '용지폭
            .SetText 27, .MaxRows, IIf(chkFontItalic(idx).Value = "0", "0", "1")    '무조건고정
            .SetText 28, .MaxRows, gblCtrlNm
        Case 2
            .SetText 1, .MaxRows, .MaxRows - 1                                      '설정순번
            .SetText 2, .MaxRows, idx                                               '항목구분
            .SetText 3, .MaxRows, txtTitle.Text                                     '항목명
            .SetText 4, .MaxRows, txtXpos.Text                                      'X1좌표
            .SetText 5, .MaxRows, txtImageWSize(0).Text                             'X2좌표
            .SetText 6, .MaxRows, txtYpos.Text                                      'Y1좌표
            .SetText 7, .MaxRows, txtImageHSize(0).Text                             'Y2좌표
            .SetText 8, .MaxRows, ""                             '폰트명
            .SetText 9, .MaxRows, "0"                             '폰트크기
            .SetText 10, .MaxRows, "0"      '폰트굵게
            .SetText 11, .MaxRows, "0"     '폰트밑줄
            .SetText 12, .MaxRows, "0"     '폰트기울게
            .SetText 13, .MaxRows, "0"                                              '폰트회전
            .SetText 14, .MaxRows, "0"                                              '바코드종류
            .SetText 15, .MaxRows, "0"                                              '바코드폭
            .SetText 16, .MaxRows, "0"                                              '바코드회전
            .SetText 17, .MaxRows, txtImageName(0).Text                                                '이미지경로
            .SetText 18, .MaxRows, "0"                                              '라인회전
            .SetText 19, .MaxRows, "0"                                              '라인두께
            .SetText 20, .MaxRows, "0"                                              '라인폭
            .SetText 21, .MaxRows, IIf(chkPrint.Value = "1", "0", "1")              '출력여부
            .SetText 22, .MaxRows, ""                             '출력값
            .SetText 23, .MaxRows, gScaleCal                                              'X좌표 보정값
            .SetText 24, .MaxRows, gScaleCal                                              'Y좌표 보정값
            .SetText 25, .MaxRows, txtPaperHSize.Text                               '용지높이
            .SetText 26, .MaxRows, txtPaperWSize.Text                               '용지폭
            .SetText 27, .MaxRows, IIf(chkIStatic.Value = "0", "0", "1")    '무조건고정
            .SetText 28, .MaxRows, gblCtrlNm
        Case 3
            .SetText 1, .MaxRows, .MaxRows - 1                                      '설정순번
            .SetText 2, .MaxRows, idx                                               '항목구분
            .SetText 3, .MaxRows, txtTitle.Text                                     '항목명
            .SetText 4, .MaxRows, txtXpos.Text                                      'X1좌표
            .SetText 5, .MaxRows, txtImageWSize(1).Text                             'X2좌표
            .SetText 6, .MaxRows, txtYpos.Text                                      'Y1좌표
            .SetText 7, .MaxRows, txtImageHSize(1).Text                             'Y2좌표
            .SetText 8, .MaxRows, ""                             '폰트명
            .SetText 9, .MaxRows, "0"                             '폰트크기
            .SetText 10, .MaxRows, "0"      '폰트굵게
            .SetText 11, .MaxRows, "0"     '폰트밑줄
            .SetText 12, .MaxRows, "0"     '폰트기울게
            .SetText 13, .MaxRows, "0"                                              '폰트회전
            .SetText 14, .MaxRows, "0"                                              '바코드종류
            .SetText 15, .MaxRows, "0"                                              '바코드폭
            .SetText 16, .MaxRows, "0"                                              '바코드회전
            .SetText 17, .MaxRows, txtImageName(1).Text                                                '이미지경로
            .SetText 18, .MaxRows, "0"                                              '라인회전
            .SetText 19, .MaxRows, "0"                                              '라인두께
            .SetText 20, .MaxRows, "0"                                              '라인폭
            .SetText 21, .MaxRows, IIf(chkPrint.Value = "1", "0", "1")              '출력여부
            .SetText 22, .MaxRows, ""                             '출력값
            .SetText 23, .MaxRows, gScaleCal                                              'X좌표 보정값
            .SetText 24, .MaxRows, gScaleCal                                              'Y좌표 보정값
            .SetText 25, .MaxRows, txtPaperHSize.Text                               '용지높이
            .SetText 26, .MaxRows, txtPaperWSize.Text                               '용지폭
            .SetText 27, .MaxRows, IIf(chkIStatic.Value = "0", "0", "1")    '무조건고정
            .SetText 28, .MaxRows, gblCtrlNm
        
        Case 4
            .SetText 1, .MaxRows, .MaxRows - 1                                      '설정순번
            .SetText 2, .MaxRows, idx                                               '항목구분
            .SetText 3, .MaxRows, txtTitle.Text                                     '항목명
            .SetText 4, .MaxRows, txtXpos.Text                                      'X1좌표
            .SetText 5, .MaxRows, txtBarWSize.Text                             'X2좌표
            .SetText 6, .MaxRows, txtYpos.Text                                      'Y1좌표
            .SetText 7, .MaxRows, txtBarHSize.Text                             'Y2좌표
            .SetText 8, .MaxRows, ""                             '폰트명
            .SetText 9, .MaxRows, "0"                             '폰트크기
            .SetText 10, .MaxRows, "0"      '폰트굵게
            .SetText 11, .MaxRows, "0"     '폰트밑줄
            .SetText 12, .MaxRows, "0"     '폰트기울게
            .SetText 13, .MaxRows, "0"                                              '폰트회전
            .SetText 14, .MaxRows, cboBarType.ListIndex                                              '바코드종류
            .SetText 15, .MaxRows, "0" 'txtBarDevide.Text                                              '바코드폭
            .SetText 16, .MaxRows, IIf(chkBarRotate.Value = "0", 0, 2)                                               '바코드회전
            .SetText 17, .MaxRows, ""                                                '이미지경로
            .SetText 18, .MaxRows, "0"                                              '라인회전
            .SetText 19, .MaxRows, "0"                                              '라인두께
            .SetText 20, .MaxRows, "0"                                              '라인폭
            .SetText 21, .MaxRows, IIf(chkPrint.Value = "1", "0", "1")              '출력여부
            .SetText 22, .MaxRows, Trim(txtBarData.Text)                              '출력값
            .SetText 23, .MaxRows, gScaleCal                                              'X좌표 보정값
            .SetText 24, .MaxRows, gScaleCal                                              'Y좌표 보정값
            .SetText 25, .MaxRows, txtPaperHSize.Text                               '용지높이
            .SetText 26, .MaxRows, txtPaperWSize.Text                               '용지폭
            .SetText 27, .MaxRows, IIf(chkIStatic.Value = "0", "0", "1")    '무조건고정
            .SetText 28, .MaxRows, strBarImgName    'Tag
            .SetText 28, .MaxRows, gblCtrlNm
        
        Case 5
            .SetText 1, .MaxRows, .MaxRows - 1                                      '설정순번
            .SetText 2, .MaxRows, idx                                               '항목구분
            .SetText 3, .MaxRows, txtTitle.Text                                     '항목명
            If chkLineRotate.Value = "0" Then
                .SetText 4, .MaxRows, txtXpos.Text                                      'X1좌표
                .SetText 5, .MaxRows, txtLineWSize.Text                             'X2좌표
                .SetText 6, .MaxRows, txtYpos.Text                                      'Y1좌표
                .SetText 7, .MaxRows, txtYpos.Text                             'Y2좌표
            Else
                .SetText 4, .MaxRows, txtXpos.Text                                      'X1좌표
                .SetText 5, .MaxRows, txtXpos.Text                             'X2좌표
                .SetText 6, .MaxRows, txtYpos.Text                                      'Y1좌표
                .SetText 7, .MaxRows, txtLineWSize.Text                             'Y2좌표
            End If
            .SetText 8, .MaxRows, ""                             '폰트명
            .SetText 9, .MaxRows, "1"                             '폰트크기
            .SetText 10, .MaxRows, "0"      '폰트굵게
            .SetText 11, .MaxRows, "0"     '폰트밑줄
            .SetText 12, .MaxRows, "0"     '폰트기울게
            .SetText 13, .MaxRows, "0"                                              '폰트회전
            .SetText 14, .MaxRows, "0"                                              '바코드종류
            .SetText 15, .MaxRows, "0"                                              '바코드폭
            .SetText 16, .MaxRows, "0"                                              '바코드회전
            .SetText 17, .MaxRows, ""                                               '이미지경로
            .SetText 18, .MaxRows, IIf(chkLineRotate.Value = "0", "0", "1")         '라인회전
            .SetText 19, .MaxRows, txtLineHSize.Text                                '라인두께
            .SetText 20, .MaxRows, txtLineWSize.Text                                '라인폭
            .SetText 21, .MaxRows, IIf(chkPrint.Value = "1", "0", "1")              '출력여부
            .SetText 22, .MaxRows, ""                                               '출력값
            .SetText 23, .MaxRows, gScaleCal                                        'X좌표 보정값
            .SetText 24, .MaxRows, gScaleCal                                        'Y좌표 보정값
            .SetText 25, .MaxRows, txtPaperHSize.Text                               '용지높이
            .SetText 26, .MaxRows, txtPaperWSize.Text                               '용지폭
            .SetText 27, .MaxRows, IIf(chkIStatic.Value = "0", "0", "1")    '무조건고정
            .SetText 28, .MaxRows, Trim(txtTag.Text)     'Tag
            .SetText 28, .MaxRows, gblCtrlNm
        
        End Select
        
'        .ColWidth(-1) = 5
    End With
    
End Sub

' 오프젝트를 생성시킨다.
Private Function objMake() As String
    Dim obj                 As Object
    Dim ClsEventObject      As ClassEventObject
    
    Set ClsEventObject = New ClassEventObject

    objMake = "0"
    
    Select Case sstType.Tab
    Case 0  'Static Label
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSLabel, txtTag.Text)
        If Not obj Is Nothing Then
            obj.Tag = txtTitle.Text
            obj.AutoSize = True
            obj.BackColor = vbWhite
            obj.Font = txtFontName(sstType.Tab).Text
            obj.FontSize = Round(txtFontSize(sstType.Tab).Text * gDevide, 0)
            obj.FontBold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
            obj.FontItalic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
            obj.FontUnderline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.Caption = txtContent(sstType.Tab).Text
            obj.DataMember = chkTStatic.Value                       '-- 무조건고정
            obj.DataField = IIf(chkPrint.Value = "1", "0", "1")     '-- 출력안함
            obj.MousePointer = 5
            
        Else
            Set ClsEventObject = Nothing
            If MsgBox("동일한 항목명은 사용할 수 없습니다." & vbNewLine & "종료하시겠습니까?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
                objMake = txtTag.Text & "_EDIT"
                Exit Function
            End If
        End If
    Case 1  'Dynamic Label
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDLabel, txtTag.Text)
        If Not obj Is Nothing Then
            obj.Tag = txtTitle.Text
            obj.AutoSize = True
            obj.BackColor = vbWhite
            obj.Font = txtFontName(sstType.Tab).Text
            obj.FontSize = Round(txtFontSize(sstType.Tab).Text * gDevide, 0)
            obj.FontBold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
            obj.FontItalic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
            obj.FontUnderline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.Caption = txtContent(sstType.Tab).Text
            obj.DataMember = IIf(chkPrint.Value = "1", "0", "1")   '-- 출력안함
            obj.MousePointer = 5
            
'            With txtText
            picPaint.CurrentX = obj.Left
            picPaint.CurrentY = obj.Top
            picPaint.Font = obj.Font
            picPaint.ForeColor = lblForeColor.BackColor
            picPaint.Print obj.Caption
            txtText.Visible = False
            SetImageBuffer
 ' End With
        Else
            Set ClsEventObject = Nothing
            If MsgBox(txtTag.Text & " 항목명은 사용할 수 없습니다." & vbNewLine & "항목명을 변경하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
                objMake = txtTag.Text & "_EDIT"
                Exit Function
            End If
        End If
    Case 2 'Static Image
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSImage, txtTag.Text)
        If Not obj Is Nothing Then
            If Dir(txtImageName(0).Text) = "" Then
                obj.Picture = LoadPicture(App.Path & "\" & gImage & "noimage.bmp")
            Else
                obj.Picture = LoadPicture(txtImageName(0).Text)
            End If
            obj.Tag = txtTitle.Text
            obj.DataMember = txtImageName(0).Text   '-- 이미지경로
            obj.Stretch = True
            obj.Width = Round(txtImageWSize(0).Text * gDevide, 0)
            obj.Height = Round(txtImageHSize(0).Text * gDevide, 0)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.ToolTipText = chkIStatic.Value      '-- 무조건고정
            obj.DataField = IIf(chkPrint.Value = "1", "0", "1")   '-- 출력안함
            obj.MousePointer = 5
        Else
'            MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
'            Set ClsEventObject = Nothing
'            Exit Function
            Set ClsEventObject = Nothing
            If MsgBox(txtTag.Text & " 항목명은 사용할 수 없습니다." & vbNewLine & "항목명을 변경하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
                objMake = txtTag.Text & "_EDIT"
                Exit Function
            End If
        
        End If
    Case 3 'Dynamic Image
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDImage, txtTag.Text)
        If Not obj Is Nothing Then
            If Dir(txtImageName(1).Text) = "" Then
                obj.Picture = LoadPicture(App.Path & "\" & gImage & "noimage.bmp")
            Else
                obj.Picture = LoadPicture(txtImageName(1).Text)
            End If
            obj.Tag = txtTitle.Text
            obj.DataMember = txtImageName(1).Text       '-- 이미지경로
            obj.Stretch = True
            obj.Width = Round(txtImageWSize(1).Text * gDevide, 0)
            obj.Height = Round(txtImageHSize(1).Text * gDevide, 0)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.DataField = IIf(chkPrint.Value = "1", "0", "1")   '-- 출력안함
            obj.MousePointer = 5
        Else
'            MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
'            Set ClsEventObject = Nothing
'            Exit Function
            Set ClsEventObject = Nothing
            If MsgBox(txtTag.Text & " 항목명은 사용할 수 없습니다." & vbNewLine & "항목명을 변경하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
                objMake = txtTag.Text & "_EDIT"
                Exit Function
            End If
        
        End If

    Case 4 'Barcode
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, txtTag.Text)
        If Not obj Is Nothing Then
            obj.Tag = txtTitle.Text
            obj.Caption = txtBarData.Text
            obj.Style = cboBarType.ListIndex
            obj.Alignment = bcALeft
            obj.BarWidth = 0
            obj.Width = Round(txtBarWSize.Text * gDevide, 0)
            obj.Height = Round(txtBarHSize.Text * gDevide, 0)
            obj.Top = Round(txtYpos.Text * gDevide, 0)
            obj.Left = Round(txtXpos.Text * gDevide, 0)
            obj.Direction = IIf(chkBarRotate.Value = "0", 0, 2)
            obj.Visible = False
'            obj.Visible = True
        
            Set obj.Container = Picture1
            m_ColCommandButton.Add ClsEventObject
            Set ClsEventObject = Nothing
            
            '== 바코드를 이미지 형태로 올리기 ===================================================================
            If intMode = 0 Then '==== Mode Set [0:로드,1:적용,2:이동,3:생성]
                If strBarImgName = "" Then
                    'strBarImgName = txtTag.Text & "_IMG1"
                    strBarImgName = txtTag.Text & "_IMG"
                Else
                    strBarImgName = Mid(strBarImgName, 1, Len(strBarImgName) - 1) & Right(strBarImgName, 1) + 1
                End If
            End If
            
            Set ClsEventObject = New ClassEventObject
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBImage, strBarImgName)
            If Not obj Is Nothing Then
                obj.Picture = LoadPicture(App.Path & "\" & gImage & "\barcode.bmp")
                obj.Tag = txtTitle.Text
                obj.DataMember = App.Path & "\" & gImage & "\barcode.bmp"   '-- 이미지 경로
                obj.Stretch = True
                obj.Width = Round(txtBarWSize.Text * gDevide, 0)
                obj.Height = Round(txtBarHSize.Text * gDevide, 0)
                obj.Top = Round(txtYpos.Text * gDevide, 0)
                obj.Left = Round(txtXpos.Text * gDevide, 0)
                obj.ToolTipText = cboBarType.ListIndex                      '-- 바코드 타입
                obj.DataField = IIf(chkPrint.Value = "1", "0", "1")         '-- 출력안함
                obj.MousePointer = 5
            Else
                MsgBox "동일한 항목명은 사용할 수 없습니다.[바코드 생성 오류]", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Function
            End If
            '== 바코드를 이미지 형태로 올리기 ===================================================================
        Else
            MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
            Set ClsEventObject = Nothing
            Exit Function
        End If
    Case 5  'Line
        Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLImage, txtTag.Text)
        If Not obj Is Nothing Then
            obj.Tag = txtTitle.Text
            If chkLineRotate.Value = 0 Then
                obj.Picture = LoadPicture(App.Path & "\" & gImage & "wline.jpg")
                obj.Stretch = True
                obj.Width = Round(txtLineWSize * gDevide, 0)
                obj.Height = Round(txtLineHSize * gDevide, 0)
                obj.Top = Round(txtYpos.Text * gDevide, 0)
                obj.Left = Round(txtXpos.Text * gDevide, 0)
                obj.ToolTipText = IIf(chkPrint.Value = "1", "0", "1")   '-- 출력안함
                obj.DataMember = "0"                                    '-- Rotate
                obj.MousePointer = 5
            Else
                obj.Picture = LoadPicture(App.Path & "\" & gImage & "hline.jpg")
                obj.Stretch = True
                obj.Width = Round(txtLineHSize * gDevide, 0)
                obj.Height = Round(txtLineWSize * gDevide, 0)
                obj.Top = Round(txtYpos.Text * gDevide, 0)
                obj.Left = Round(txtXpos.Text * gDevide, 0)
                obj.ToolTipText = IIf(chkPrint.Value = "1", "0", "1")   '-- 출력안함
                obj.DataMember = "1"                                    '-- Rotate
                obj.MousePointer = 5
            End If
        Else
'            MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
'            Set ClsEventObject = Nothing
'            Exit Function
            Set ClsEventObject = Nothing
            If MsgBox(txtTag.Text & " 항목명은 사용할 수 없습니다." & vbNewLine & "항목명을 변경하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
                objMake = txtTag.Text & "_EDIT"
                Exit Function
            End If
        End If
    End Select
        
    obj.Visible = True
    Set obj.Container = picPaint
    

'    txtText_DblClick
    
    m_ColCommandButton.Add ClsEventObject
    Set ClsEventObject = Nothing
    
End Function

Private Sub MakeBarImage(ByVal BarObj As Object)
    
    Picture2.Height = BarObj.Height
    Picture2.Width = BarObj.Width
    Barcod1.PrinterScaleMode = vbTwips 'Form1.ScaleMode
    Barcod1.PrinterWidth = BarObj.Width
    Barcod1.PrinterHeight = BarObj.Height
    Barcod1.PrinterTop = 0
    Barcod1.PrinterLeft = 0
    Barcod1.PrinterHDC = Picture2.hDC
    Picture2.Refresh
    Clipboard.Clear
    Clipboard.SetData Picture2.Image

'    SavePicture Picture2.Image, "C:\TEST.BMP"
    SavePicture Picture2.Image, "C:\TEST.BMP"

End Sub

Private Function findSameCtrlNm(strIdx As String, strTitle As String) As Boolean
    Dim i As Integer
    Dim strCtrlIdx  As String
    Dim strCtrlNm   As String
    
    findSameCtrlNm = False
    With spdList
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2: strCtrlIdx = Trim(.Text)
            .Col = 3: strCtrlNm = Trim(.Text)
            If strIdx = strCtrlIdx And strTitle = strCtrlNm Then
                findSameCtrlNm = True
                Exit For
            End If
        Next
    End With
    
End Function

Private Sub objNewMake()
    Dim obj                 As Object
    Dim i                   As Integer
    Dim ClsEventObject      As ClassEventObject
    
    '-- 유효성 검사 [항목명]
    If Trim(txtTitle.Text) = "" Then
        MsgBox "항목명을 입력하세요.", vbInformation, Me.Caption
        txtTitle.SetFocus
        Exit Sub
    End If
    '-- 유효성 검사 [X 좌표명]
    If Trim(txtXpos.Text) = "" Then
        MsgBox "X좌표를 입력하세요.", vbInformation, Me.Caption
        txtXpos.SetFocus
        Exit Sub
    End If
    '-- 유효성 검사 [X 좌표]
    If Not IsNumeric(Trim(txtXpos.Text)) Then
        MsgBox "X좌표는 숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
        txtXpos.SetFocus
        Exit Sub
    End If
    '-- 유효성 검사 [Y 좌표명]
    If Trim(txtYpos.Text) = "" Then
        MsgBox "Y좌표를 입력하세요.", vbInformation, Me.Caption
        txtYpos.SetFocus
        Exit Sub
    End If
    '-- 유효성 검사 [Y 좌표]
    If Not IsNumeric(Trim(txtYpos.Text)) Then
        MsgBox "Y좌표는 숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
        txtYpos.SetFocus
        Exit Sub
    End If
            
    Select Case sstType.Tab
        Case 0 '## Static Label ##
            '-- 유효성 검사 [폰트명]
            If Trim(txtFontName(0).Text) = "" Or Trim(txtFontSize(0).Text) = "" Then
                MsgBox "Font를 선택하세요.", vbInformation, Me.Caption
                Call cmdFont_Click(0)
                Exit Sub
            End If
            '-- 유효성 검사 [폰트사이즈]
            If Not IsNumeric(Trim(txtFontSize(0).Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
                txtFontSize(0).SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [텍스트]
            If Trim(txtContent(0).Text) = "" Then
                MsgBox "Text를 입력하세요.", vbInformation, Me.Caption
                txtContent(0).SetFocus
                Exit Sub
            End If
            
            '-- 동일명칭 체크
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Static Label 개체만들기
            gblCtrlIdx = gblCtrlIdx + 1
            gblCtrlNm = "Control_" & gblCtrlIdx
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSLabel, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSLabel, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                obj.AutoSize = True
                obj.BackColor = vbWhite
                obj.Font = txtFontName(sstType.Tab).Text
                obj.FontSize = txtFontSize(sstType.Tab).Text * gDevide
                obj.FontBold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
                obj.FontItalic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
                obj.FontUnderline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.Caption = txtContent(sstType.Tab).Text
                obj.DataMember = chkTStatic.Value              '-- 무조건고정
                obj.DataField = IIf(chkPrint.Value = "1", "0", "1")     '-- 출력안함
                obj.MousePointer = 5
                
                'obj======그리는곳
                'X , Y====좌표
                'Txt======글자
                'TxtGag===글자의 기울기
                'H========글자의 높이(1에 대한 배율)
                'W========글자의 너비(1에 대한 배율)
                'LineSpace ====줄간격(1에 대한 배율)
                
'                Call RotateControl(obj, 90)
                
'                If optSTRotate(0).Value = True Then
'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 0, 1, 1, 1)
'
'                ElseIf optSTRotate(1).Value = True Then
'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 90, 1, 1, 1)
'                ElseIf optSTRotate(2).Value = True Then
'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 180, 1, 1, 1)
'                Else
'                    Call FontStuff(Picture1, obj.Top, obj.Left, obj.Caption, 270, 1, 1, 1)
'                End If
        

                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
        
        Case 1  '## Dynamic Label ##
            '-- 유효성 검사 [폰트명]
            If Trim(txtFontName(1).Text) = "" Or Trim(txtFontSize(1).Text) = "" Then
                MsgBox "Font를 선택하세요.", vbInformation, Me.Caption
                Call cmdFont_Click(1)
                Exit Sub
            End If
            '-- 유효성 검사 [폰트사이즈]
            If Not IsNumeric(Trim(txtFontSize(1).Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
                txtFontSize(1).SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [텍스트]
            If Trim(txtContent(1).Text) = "" Then
                MsgBox "Text를 입력하세요.", vbInformation, Me.Caption
                txtContent(1).SetFocus
                Exit Sub
            End If
            
            '-- 동일명칭 체크
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Dynamic Label 개체만들기
            gblCtrlIdx = gblCtrlIdx + 1
            gblCtrlNm = "Control_" & gblCtrlIdx
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDLabel, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDLabel, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                obj.AutoSize = True
                obj.BackColor = vbWhite
                obj.Font = txtFontName(sstType.Tab).Text
                obj.FontSize = txtFontSize(sstType.Tab).Text * gDevide
                obj.FontBold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
                obj.FontItalic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
                obj.FontUnderline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.Caption = txtContent(sstType.Tab).Text
                obj.DataMember = IIf(chkPrint.Value = "1", "0", "1")   '-- 출력안함
                obj.MousePointer = 5
                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
        
        Case 2 '## Static Image ##
            '-- 유효성 검사 [이미지명]
            If Trim(txtImageName(0).Text) = "" Then
                MsgBox "이미지를 선택하세요.", vbInformation, Me.Caption
                Call cmdImage_Click(0)
                Exit Sub
            End If
            '-- 유효성 검사 [가로Size]
            If Trim(txtImageWSize(0).Text) = "" Then
                MsgBox "가로Size를 입력하세요.", vbInformation, Me.Caption
                txtImageWSize(0).SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [가로Size]
            If Not IsNumeric(Trim(txtImageWSize(0).Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbInformation, Me.Caption
                txtImageWSize(0).SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [세로Size]
            If Trim(txtImageHSize(0).Text) = "" Then
                MsgBox "세로Size를 입력하세요.", vbInformation, Me.Caption
                txtImageHSize(0).SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [세로Size]
            If Not IsNumeric(Trim(txtImageHSize(0).Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbInformation, Me.Caption
                txtImageHSize(0).SetFocus
                Exit Sub
            End If
            
            '-- 동일명칭 체크
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Static Image 개체만들기
            gblCtrlIdx = gblCtrlIdx + 1
            gblCtrlNm = "Control_" & gblCtrlIdx
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSImage, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectSImage, gblCtrlNm)
            If Not obj Is Nothing Then
                If Dir(txtImageName(0).Text) = "" Then
                    obj.Picture = LoadPicture(App.Path & "\image\noimage.bmp")
                Else
                    obj.Picture = LoadPicture(txtImageName(0).Text)
                End If
                obj.Tag = txtTitle.Text
                obj.DataMember = txtImageName(0).Text           '-- 이미지 경로
                obj.Stretch = True
                obj.Width = txtImageWSize(0).Text * gDevide
                obj.Height = txtImageHSize(0).Text * gDevide
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.MousePointer = 5
                obj.ToolTipText = chkIStatic.Value              '-- 무조건고정
                obj.DataField = IIf(chkPrint.Value = "1", "0", "1")   '-- 출력안함
                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
            
        Case 3 '## Dynamic Image ##
            '-- 유효성 검사 [이미지명]
            If Trim(txtImageName(1).Text) = "" Then
                MsgBox "이미지를 선택하세요.", vbInformation, Me.Caption
                Call cmdImage_Click(1)
                Exit Sub
            End If
            '-- 유효성 검사 [가로Size]
            If Trim(txtImageWSize(1).Text) = "" Then
                MsgBox "가로Size를 입력하세요.", vbInformation, Me.Caption
                txtImageWSize(1).SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [가로Size]
            If Not IsNumeric(Trim(txtImageWSize(1).Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbInformation, Me.Caption
                txtImageWSize(1).SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [세로Size]
            If Trim(txtImageHSize(1).Text) = "" Then
                MsgBox "세로Size를 입력하세요.", vbInformation, Me.Caption
                txtImageHSize(1).SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [세로Size]
            If Not IsNumeric(Trim(txtImageHSize(1).Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbInformation, Me.Caption
                txtImageHSize(1).SetFocus
                Exit Sub
            End If
            
            '-- 동일명칭 체크
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Dynamic Image 개체만들기
            gblCtrlIdx = gblCtrlIdx + 1
            gblCtrlNm = "Control_" & gblCtrlIdx
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDImage, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectDImage, gblCtrlNm)
            If Not obj Is Nothing Then
                If Dir(txtImageName(1).Text) = "" Then
                    obj.Picture = LoadPicture(App.Path & "\image\noimage.bmp")
                Else
                    obj.Picture = LoadPicture(txtImageName(1).Text)
                End If
                obj.Tag = txtTitle.Text
                obj.DataMember = txtImageName(1).Text           '-- 이미지 경로
                obj.Stretch = True
                obj.Width = txtImageWSize(1).Text * gDevide
                obj.Height = txtImageHSize(1).Text * gDevide
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.DataField = IIf(chkPrint.Value = "1", "0", "1")   '-- 출력안함
                obj.MousePointer = 5
                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
    
        Case 4  '## Barcode ##
            '-- 유효성 검사 [길이Size]
            If Trim(txtBarWSize.Text) = "" Then
                MsgBox "길이Size를 입력하세요.", vbInformation, Me.Caption
                txtBarWSize.SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [길이Size]
            If Not IsNumeric(Trim(txtBarWSize.Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbInformation, Me.Caption
                txtBarWSize.SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [높이Size]
            If Trim(txtBarHSize.Text) = "" Then
                MsgBox "높이Size를 입력하세요.", vbInformation, Me.Caption
                txtBarHSize.SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [높이Size]
            If Not IsNumeric(Trim(txtBarHSize.Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbInformation, Me.Caption
                txtBarHSize.SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [높이Size]
            If Trim(txtBarData.Text) = "" Then
                MsgBox "Data를 입력하세요.", vbInformation, Me.Caption
                txtBarData.SetFocus
                Exit Sub
            End If
            
            '-- 동일명칭 체크
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Barcode 개체만들기
            gblCtrlIdx = gblCtrlIdx + 1
            gblCtrlNm = "Control_" & gblCtrlIdx
            
            Set ClsEventObject = New ClassEventObject
            'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, txtTitle.Text)
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                obj.Caption = txtBarData.Text
                obj.Style = cboBarType.ListIndex
                obj.Alignment = bcALeft
                obj.BarWidth = 0
                obj.Width = txtBarWSize.Text * gDevide
                obj.Height = txtBarHSize.Text * gDevide
                obj.Top = txtYpos.Text * gDevide
                obj.Left = txtXpos.Text * gDevide
                obj.Direction = IIf(chkBarRotate.Value = "0", 0, 2)
                'obj.DataField = IIf(chkPrint.Value = "1", "0", "1")         '-- 출력안함
                obj.Visible = False
                                
                Set obj.Container = Picture1
                m_ColCommandButton.Add ClsEventObject
                Set ClsEventObject = Nothing
                
'                If strBarImgName = "" Then
'                    strBarImgName = txtTitle.Text & "_IMG1"
'                Else
'                    strBarImgName = Mid(strBarImgName, 1, Len(strBarImgName) - 1) & Right(strBarImgName, 1) + 1
'                End If

                '-- 동일명칭 체크
                If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                    MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                    Exit Sub
                End If

                gblCtrlNm = gblCtrlNm & "_IMG"
                Call MakeSpdSaveList(obj, sstType.Tab)
                                
                '== 바코드를 이미지 형태로 올리기 ===================================================================
                'gblCtrlNm = gblCtrlNm & "_IMG"
                
                Set ClsEventObject = New ClassEventObject
                'Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBImage, strBarImgName)
                Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBImage, gblCtrlNm)
                If Not obj Is Nothing Then
                    obj.Picture = LoadPicture(App.Path & "\" & gImage & "\barcode.bmp")
                    obj.Tag = txtTitle.Text
                    obj.DataMember = App.Path & "\" & gImage & "\barcode.bmp"
                    obj.Stretch = True
                    obj.Width = txtBarWSize.Text * gDevide
                    obj.Height = txtBarHSize.Text * gDevide
                    obj.Top = txtYpos.Text * gDevide
                    obj.Left = txtXpos.Text * gDevide
                    obj.ToolTipText = cboBarType.ListIndex
                    obj.DataField = IIf(chkPrint.Value = "1", "0", "1")         '-- 출력안함
                    obj.MousePointer = 5
                Else
                    If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                    MsgBox "동일한 항목명은 사용할 수 없습니다.[바코드 생성 오류]", vbInformation, Me.Caption
                    Set ClsEventObject = Nothing
                    Exit Sub
                End If
                '== 바코드를 이미지 형태로 올리기 ===================================================================
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
    
        Case 5  '## Line ##
            '-- 유효성 검사 [선굵기]
            If Trim(txtLineHSize.Text) = "" Then
                MsgBox "선굵기를 입력하세요.", vbInformation, Me.Caption
                txtLineHSize.SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [선굵기]
            If Not IsNumeric(Trim(txtLineHSize.Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbInformation, Me.Caption
                txtLineHSize.SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [선길이]
            If Trim(txtLineWSize.Text) = "" Then
                MsgBox "선길이를 입력하세요.", vbInformation, Me.Caption
                txtLineWSize.SetFocus
                Exit Sub
            End If
            '-- 유효성 검사 [선길이]
            If Not IsNumeric(Trim(txtLineWSize.Text)) Then
                MsgBox "숫자만 입력이 가능합니다.", vbInformation, Me.Caption
                txtLineWSize.SetFocus
                Exit Sub
            End If
            
            '-- 동일명칭 체크
            If findSameCtrlNm(sstType.Tab, txtTitle.Text) Then
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '-- Line 개체만들기
            gblCtrlIdx = gblCtrlIdx + 1
            gblCtrlNm = "Control_" & gblCtrlIdx
            
            Set ClsEventObject = New ClassEventObject
            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLImage, gblCtrlNm)
            If Not obj Is Nothing Then
                obj.Tag = txtTitle.Text
                If chkLineRotate.Value = 0 Then
                    obj.Picture = LoadPicture(App.Path & "\" & gImage & "wline.jpg")
                    obj.Stretch = True
                    obj.Width = txtLineWSize * gScaleCal
                    obj.Height = txtLineHSize * gScaleCal
                    obj.Top = txtYpos.Text * gScaleCal
                    obj.Left = txtXpos.Text * gScaleCal
                    obj.DataMember = "0"
                Else
                    obj.Picture = LoadPicture(App.Path & "\" & gImage & "hline.jpg")
                    obj.Stretch = True
                    obj.Width = txtLineHSize * gScaleCal
                    obj.Height = txtLineWSize * gScaleCal
                    obj.Top = txtYpos.Text * gScaleCal
                    obj.Left = txtXpos.Text * gScaleCal
                    obj.DataMember = "1"
                End If
                obj.MousePointer = 5
                Call MakeSpdSaveList(obj, sstType.Tab)
            Else
                If gblCtrlIdx > 0 Then gblCtrlIdx = gblCtrlIdx - 1
                MsgBox "동일한 항목명은 사용할 수 없습니다.", vbInformation, Me.Caption
                Set ClsEventObject = Nothing
                Exit Sub
            End If
    End Select
        
    
'    Dim lnghNewFont As Long
'    Dim lnghOriginalFonrt As Long
'    Dim lngHeight As Long
'    Dim lngWidth As Long
'    Dim intAngle As Integer
    
    
    obj.Visible = True
    Set obj.Container = Picture1
    
    m_ColCommandButton.Add ClsEventObject
    
    Set ClsEventObject = Nothing
    
'    intAngle = 90
'    With Picture1
'        .ScaleMode = vbPixels
'        .AutoRedraw = True
'        lngHeight = .TextHeight(obj)
'        lngWidth = 0
'
'        With .Font
'            lnghNewFont = CreateFont(lngHeight, lngWidth, intAngle * 10, intAngle * 10, .Weight, .Italic, .Underline, .Strikethrough, .Charset, 0, 0, 0, 0, .Name)
'        End With
'        lnghOriginalFonrt = SelectObject(.hdc, lnghNewFont)
'        .CurrentX = obj.Left
'        .CurrentY = obj.Top
'        Picture1.Print obj
'
'        lnghNewFont = SelectObject(.hdc, lnghOriginalFonrt)
'        .AutoRedraw = False
'    End With
'    DeleteObject lnghNewFont
'    'obj.Visible = False
        
    
End Sub

Private Sub objSet()
    Dim strNm As String

    Select Case sstType.Tab
    Case 0  'Static Label
            Me.Controls(txtTag.Text).AutoSize = True
            Me.Controls(txtTag.Text).BackColor = vbWhite
            Me.Controls(txtTag.Text).Font = txtFontName(sstType.Tab).Text
            Me.Controls(txtTag.Text).FontSize = txtFontSize(sstType.Tab).Text * gDevide
            Me.Controls(txtTag.Text).FontBold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).FontItalic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).FontUnderline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).Top = txtYpos.Text * gDevide
            Me.Controls(txtTag.Text).Left = txtXpos.Text * gDevide
            Me.Controls(txtTag.Text).Caption = txtContent(sstType.Tab).Text
            Me.Controls(txtTag.Text).DataMember = chkTStatic.Value
            Me.Controls(txtTag.Text).DataField = IIf(chkPrint.Value = "1", "0", "1")    '-- 출력안함
            
    Case 1  'Dynamic Label
            Me.Controls(txtTag.Text).AutoSize = True
            Me.Controls(txtTag.Text).BackColor = vbWhite
            Me.Controls(txtTag.Text).Font = txtFontName(sstType.Tab).Text
            Me.Controls(txtTag.Text).FontSize = txtFontSize(sstType.Tab).Text * gDevide
            Me.Controls(txtTag.Text).FontBold = IIf(chkFontBold(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).FontItalic = IIf(chkFontItalic(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).FontUnderline = IIf(chkFontUnder(sstType.Tab).Value = 1, True, False)
            Me.Controls(txtTag.Text).Top = txtYpos.Text * gDevide
            Me.Controls(txtTag.Text).Left = txtXpos.Text * gDevide
            Me.Controls(txtTag.Text).Caption = txtContent(sstType.Tab).Text
            Me.Controls(txtTag.Text).DataMember = IIf(chkPrint.Value = "1", "0", "1")          '-- 출력안함
    
    Case 2 'Static Image
            Me.Controls(txtTag.Text).Width = txtImageWSize(0).Text * gDevide
            Me.Controls(txtTag.Text).Height = txtImageHSize(0).Text * gDevide
            Me.Controls(txtTag.Text).Top = txtYpos.Text * gDevide
            Me.Controls(txtTag.Text).Left = txtXpos.Text * gDevide
            If Dir(txtImageName(0).Text) = "" Then
                Me.Controls(txtTag.Text).Picture = LoadPicture(App.Path & "\" & gImage & "noimage.bmp")
            Else
                Me.Controls(txtTag.Text).Picture = LoadPicture(txtImageName(0).Text)
            End If
            
            Me.Controls(txtTag.Text).DataMember = txtImageName(0).Text   '-- 이미지경로
            
            Me.Controls(txtTag.Text).DataField = IIf(chkPrint.Value = "1", "0", "1")    '-- 출력안함
            
    Case 3 'Dynamic Image
            Me.Controls(txtTag.Text).Width = txtImageWSize(1).Text * gDevide
            Me.Controls(txtTag.Text).Height = txtImageHSize(1).Text * gDevide
            Me.Controls(txtTag.Text).Top = txtYpos.Text * gDevide
            Me.Controls(txtTag.Text).Left = txtXpos.Text * gDevide
        
            If Dir(txtImageName(1).Text) = "" Then
                Me.Controls(txtTag.Text).Picture = LoadPicture(App.Path & "\" & gImage & "noimage.bmp")
            Else
                Me.Controls(txtTag.Text).Picture = LoadPicture(txtImageName(1).Text)
            End If

            Me.Controls(txtTag.Text).DataMember = txtImageName(1).Text   '-- 이미지경로
            Me.Controls(txtTag.Text).DataField = IIf(chkPrint.Value = "1", "0", "1")    '-- 출력안함
        
    Case 4  'Barcode Label
            '-- 바코드 이미지 적용
            strNm = txtTag.Text
            Me.Controls(strNm).Width = txtBarWSize.Text * gDevide
            Me.Controls(strNm).Height = txtBarHSize.Text * gDevide
            Me.Controls(strNm).Top = txtYpos.Text * gDevide
            Me.Controls(strNm).Left = txtXpos.Text * gDevide
            Me.Controls(strNm).Picture = LoadPicture(App.Path & "\" & gImage & "barcode.bmp")
            Me.Controls(strNm).ToolTipText = cboBarType.ListIndex           '-- 바코드 타입
            Me.Controls(strNm).DataField = IIf(chkPrint.Value = "1", "0", "1")    '-- 출력안함
            
            '-- 바코드 적용
            strNm = Mid(Trim(txtTag.Text), 1, InStr(Trim(txtTag.Text), "_IMG") - 1)
            Me.Controls(strNm).Caption = txtBarData.Text
            Me.Controls(strNm).Style = cboBarType.ListIndex
            Me.Controls(strNm).Alignment = bcALeft
            Me.Controls(strNm).Width = txtBarWSize.Text * gDevide
            Me.Controls(strNm).Height = txtBarHSize.Text * gDevide
            Me.Controls(strNm).Top = txtYpos.Text * gDevide
            Me.Controls(strNm).Left = txtXpos.Text * gDevide
            Me.Controls(strNm).Direction = IIf(chkBarRotate.Value = "0", 0, 2)
            
            
    Case 5  'Line Image
            If chkLineRotate.Value = 0 Then
                Me.Controls(txtTag.Text).Width = txtLineWSize * gDevide
                Me.Controls(txtTag.Text).Height = txtLineHSize * gDevide
                Me.Controls(txtTag.Text).Top = txtYpos.Text * gDevide
                Me.Controls(txtTag.Text).Left = txtXpos.Text * gDevide
            Else
                Me.Controls(txtTag.Text).Width = txtLineHSize * gDevide
                Me.Controls(txtTag.Text).Height = txtLineWSize * gDevide
                Me.Controls(txtTag.Text).Top = txtYpos.Text * gDevide
                Me.Controls(txtTag.Text).Left = txtXpos.Text * gDevide
            End If
            Me.Controls(txtTag.Text).ToolTipText = IIf(chkPrint.Value = "1", "0", "1")   '-- 출력안함
            
    End Select
    
    Call SetLayout(sstType.Tab)
        
End Sub



Private Sub cmdImageDevSet_Click(Index As Integer)

    If Trim(txtImageDevide(Index).Text) = "" And IsNumeric(txtImageDevide(Index).Text) Then
        MsgBox "이미지 배율을 확인하세요", vbOKOnly + vbInformation, Me.Caption
        txtImageDevide(Index).SetFocus
        Exit Sub
    End If
    
    If Trim(txtImageWSize(Index).Text) = "" And Trim(txtImageHSize(Index).Text) = "" And IsNumeric(txtImageWSize(Index).Text) And IsNumeric(txtImageHSize(Index).Text) Then
        MsgBox "이미지 사이즈를 확인하세요", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    Else
        txtImageWSize(Index).Text = Round(txtImageWSize(Index + 2).Text * (txtImageDevide(Index).Text / 100), 0)
        txtImageHSize(Index).Text = Round(txtImageHSize(Index + 2).Text * (txtImageDevide(Index).Text / 100), 0)
    End If
        
End Sub

' 동적 컨트롤 생성
Private Sub cmdMake_Click()
    
    '-- Mode Set [생성]
    intMode = 3
    
    Call objNewMake
        
End Sub


Private Sub objMove(Index)
    Dim intRow          As Integer
    Dim strObjType      As Variant
    Dim strObjName      As Variant
    Dim strObjRotate    As Variant
    
    With spdList
        Select Case Index
        Case 0      'left   - x1 좌표
            For intRow = 1 To .MaxRows
                .Row = intRow
                Call .GetText(2, intRow, strObjType)
                Call .GetText(29, intRow, strObjName)
                
                '-- 선택이동
                If chkChoice.Value = "1" Then
                    If strObjName = Trim(txtTag.Text) Then
                        If strObjType = 5 Then
                            If chkDetail.Value = 1 Then
                                .Col = 5: .Text = Trim(.Text) - 1
                                .Col = 4: .Text = Trim(.Text) - 1
                            Else
                                .Col = 5: .Text = Trim(.Text) - 5
                                .Col = 4: .Text = Trim(.Text) - 5
                            End If
                        Else
                            If chkDetail.Value = 1 Then
                                .Col = 4: .Text = Trim(.Text) - 1
                            Else
                                .Col = 4: .Text = Trim(.Text) - 5
                            End If
                        End If
                        '-- 라인회전[strObjRotate]이 "1" 이면 좌/우 라인이다
                        '-- XI,X2를 같이 변경해 주어야 한다.
                        'Call .GetText(18, intRow, strObjRotate)
                        Me.Controls(strObjName).Left = .Text * gDevide
                    
                    End If
                Else
                    If strObjType = 5 Then
                        If chkDetail.Value = 1 Then
                            .Col = 5: .Text = Trim(.Text) - 1
                            .Col = 4: .Text = Trim(.Text) - 1
                        Else
                            .Col = 5: .Text = Trim(.Text) - 5
                            .Col = 4: .Text = Trim(.Text) - 5
                        End If
                    Else
                        If chkDetail.Value = 1 Then
                            .Col = 4: .Text = Trim(.Text) - 1
                        Else
                            .Col = 4: .Text = Trim(.Text) - 5
                        End If
                    End If
                    '-- 라인회전[strObjRotate]이 "1" 이면 좌/우 라인이다
                    '-- XI,X2를 같이 변경해 주어야 한다.
                    'Call .GetText(18, intRow, strObjRotate)
                    Me.Controls(strObjName).Left = .Text * gDevide
                End If
            Next
        Case 1      'right  + x1 좌표
            For intRow = 1 To .MaxRows
                .Row = intRow
                Call .GetText(2, intRow, strObjType)
                Call .GetText(29, intRow, strObjName)
                'Call .GetText(18, intRow, strObjRotate)
                
                '-- 선택이동
                If chkChoice.Value = "1" Then
                    If strObjName = Trim(txtTag.Text) Then
                        If strObjType = 5 Then
                            If chkDetail.Value = 1 Then
                                .Col = 5: .Text = Trim(.Text) + 1
                                .Col = 4: .Text = Trim(.Text) + 1
                            Else
                                .Col = 5: .Text = Trim(.Text) + 5
                                .Col = 4: .Text = Trim(.Text) + 5
                            End If
                        Else
                            If chkDetail.Value = 1 Then
                                .Col = 4: .Text = Trim(.Text) + 1
                            Else
                                .Col = 4: .Text = Trim(.Text) + 5
                            End If
                        End If
                        '-- 라인회전[strObjRotate]이 "1" 이면 좌/우 라인이다
                        '-- XI,X2를 같이 변경해 주어야 한다.
                        'Call .GetText(18, intRow, strObjRotate)
                        Me.Controls(strObjName).Left = .Text * gDevide
                    
                    End If
                Else
                    If strObjType = 5 Then
                        If chkDetail.Value = 1 Then
                            .Col = 5: .Text = Trim(.Text) + 1
                            .Col = 4: .Text = Trim(.Text) + 1
                        Else
                            .Col = 5: .Text = Trim(.Text) + 5
                            .Col = 4: .Text = Trim(.Text) + 5
                        End If
                    Else
                        If chkDetail.Value = 1 Then
                            .Col = 4: .Text = Trim(.Text) + 1
                        Else
                            .Col = 4: .Text = Trim(.Text) + 5
                        End If
                    End If
                    '-- 라인회전[strObjRotate]이 "1" 이면 좌/우 라인이다
                    '-- XI,X2를 같이 변경해 주어야 한다.
                    Call .GetText(18, intRow, strObjRotate)
                    Me.Controls(strObjName).Left = .Text * gDevide
                End If
            Next
        Case 2      'top    - y1 좌표
            For intRow = 1 To .MaxRows
                .Row = intRow
                Call .GetText(2, intRow, strObjType)
                Call .GetText(29, intRow, strObjName)
                
                '-- 선택이동
                If chkChoice.Value = "1" Then
                    If strObjName = Trim(txtTag.Text) Then
                        If strObjType = 5 Then
                            If chkDetail.Value = 1 Then
                                .Col = 7: .Text = Trim(.Text) - 1
                                .Col = 6: .Text = Trim(.Text) - 1
                            Else
                                .Col = 7: .Text = Trim(.Text) - 5
                                .Col = 6: .Text = Trim(.Text) - 5
                            End If
                        Else
                            If chkDetail.Value = 1 Then
                                .Col = 6: .Text = Trim(.Text) - 1
                            Else
                                .Col = 6: .Text = Trim(.Text) - 5
                            End If
                        End If
                        Me.Controls(strObjName).Top = .Text * gDevide
                    End If
                Else
                    If strObjType = 5 Then
                        If chkDetail.Value = 1 Then
                            .Col = 7: .Text = Trim(.Text) - 1
                            .Col = 6: .Text = Trim(.Text) - 1
                        Else
                            .Col = 7: .Text = Trim(.Text) - 5
                            .Col = 6: .Text = Trim(.Text) - 5
                        End If
                    Else
                        If chkDetail.Value = 1 Then
                            .Col = 6: .Text = Trim(.Text) - 1
                        Else
                            .Col = 6: .Text = Trim(.Text) - 5
                        End If
                    End If
                    Me.Controls(strObjName).Top = .Text * gDevide
                End If
            Next
        Case 3      'bottom + y1 좌표
            For intRow = 1 To .MaxRows
                .Row = intRow
                Call .GetText(2, intRow, strObjType)
                Call .GetText(29, intRow, strObjName)

                '-- 선택이동
                If chkChoice.Value = "1" Then
                    If strObjName = Trim(txtTag.Text) Then
                        If strObjType = 5 Then
                            If chkDetail.Value = 1 Then
                                .Col = 7: .Text = Trim(.Text) + 1
                                .Col = 6: .Text = Trim(.Text) + 1
                            Else
                                .Col = 7: .Text = Trim(.Text) + 5
                                .Col = 6: .Text = Trim(.Text) + 5
                            End If
                        Else
                            If chkDetail.Value = 1 Then
                                .Col = 6: .Text = Trim(.Text) + 1
                            Else
                                .Col = 6: .Text = Trim(.Text) + 5
                            End If
                        End If
                        Me.Controls(strObjName).Top = .Text * gDevide
                    End If
                Else
                    If strObjType = 5 Then
                        If chkDetail.Value = 1 Then
                            .Col = 7: .Text = Trim(.Text) + 1
                            .Col = 6: .Text = Trim(.Text) + 1
                        Else
                            .Col = 7: .Text = Trim(.Text) + 5
                            .Col = 6: .Text = Trim(.Text) + 5
                        End If
                    Else
                        If chkDetail.Value = 1 Then
                            .Col = 6: .Text = Trim(.Text) + 1
                        Else
                            .Col = 6: .Text = Trim(.Text) + 5
                        End If
                    End If
                    Me.Controls(strObjName).Top = .Text * gDevide
                End If
            Next
        Case 4
            '-- X1,Y1 좌표설정
            For intRow = 1 To .MaxRows
                .Row = intRow
                Call .GetText(2, intRow, strObjType)
                Call .GetText(29, intRow, strObjName)
                '
                If strObjType = sstType.Tab And strObjName = Trim(txtTag.Text) Then
                    .Col = 4: .Text = Trim(txtXpos.Text)
                    Me.Controls(strObjName).Left = .Text * gDevide
                    .Col = 6: .Text = Trim(txtYpos.Text)
                    Me.Controls(strObjName).Top = .Text * gDevide
                    Exit For
                End If
            Next
        End Select
    End With

End Sub

Private Sub cmdMove_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    '-- Mode Set [이동]
    intMode = 2
    
    Call objMove(Index)
    
    If Index < 4 Then
        intMoveIdx = Index
        
        If chkContinue.Value = 1 Then
            tmrMove.Interval = 100
            tmrMove.Enabled = True
            DoEvents
        Else
            tmrMove.Enabled = False
        End If
    End If
    
End Sub

Private Sub cmdMove_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    tmrMove.Enabled = False

End Sub

Private Sub cmdPrint_Click()
    Dim prtSelectPrinter As Printer
    Dim boolPrinter_Select_Fales As Boolean
    Dim Buffer As String
    Dim aryPrinter() As String
    Dim strBuffer As String
 
'Printer 개체를 이용해 인쇄물을 작성하실 때에는 다음의 사항을 기억하여 주십시오.
'
'PaperSize 는 Printer Driver에 따라 다르지만 기본적으로 A4 용지로 지정되어 있습니다.
'용지의 크기를 사용자 정의의 크기로 지정하기 위하여 값을 256 으로 지정할 수 있지만
'용지의 크기를 지정하는 것은 무의미합니다. 게다가 256으로 변경할 때 오류를 리터하는
'드라이버들도 더러 있기 때문입니다.
'용지의 크기를 지정할 필요는 없으며 인쇄물의 크기만 신경쓰시면 되겠습니다.
'
'님께서 적어놓은 코드를 보자면 가로 190, 세로 134 mm 의 용지에 맞게 출력을 하실려고
'하는 것 같습니다.
'이럴 경우 용지의 크기는 190 * 134 보다 작지만 않다면 어떤 용지규격으로 셋팅해도 관계
'없습니다. 이럴 경우에는 그냥 A4 로 셋팅하셔도 됩니다.
'Printer의 Width속성과 Height속성은 Twip 단위로 되어 있으며 현재 인쇄가능한 인쇄물의
'테두리(한계, Boundary)개념으로 생각하시는 게 좋을 듯 합니다.

'인쇄할 때 가장 중요한 것은 ScaleMode, Scale, ScaleWidth, ScaleHeight 입니다.
'
'mm 단위측정값을 기준으로 출력하시고자 한다면 ScaleMode 를  6 으로 지정하시면 됩니다.
'주의할 점은 용지를 A4로, ScaleMode를 6 으로 셋팅한 후에
'Printer.Line (0, 0)-(210, 297), , B
'위의 문을 실행했을 경우 우측과 하단의 테두리는 범위가 초과하여 출력이 되지 않습니다.
'왜냐하면 용지의 크기는 210 * 297 이지만 프린터마다 인쇄가능영역이라는 게 존재합니다.
'잉크젯인 경우에는 레이저젯에 비해 같은 용지에 대해 인쇄가능영역이 작습니다.
'그래서 ScaleMode 를 6으로 했을 때 ScaleWidth 나 ScaleHeight의 값을 보면 210 또는 297 보다
'작은 값으로 되어 있다는 것을 알 수 있습니다.
'이런 부분들을 고려하여 인쇄물을 작성해 보시기 바랍니다.
'그럼 즐프~~하세요.

 

''    '============== 이미지 출력 방식 ==========================================================
''    Picture1.AutoRedraw = True
''    SendMessage Picture1.hwnd, WM_PAINT, Picture1.hDC, 0
''    'SendMessage Picture1.hwnd, WM_PRINT, Picture1.hDC, PRF_CHILDREN Or PRF_CLIENT Or PRF_OWNED
''    Printer.PaintPicture Picture1.Image, 0, 0, Picture1.Width, Picture1.Height
''    Printer.EndDoc
''    SavePicture Picture1.Image, "C:\TEST.BMP"
    
''    '============== 이미지 출력 방식 ==========================================================
    
'Exit Sub

    Dim intRow As Integer
    Dim intCol As Integer
    Dim intCnt As Integer
    Dim strX1, strX2, strY1, strY2 As String
    Dim strFont As String
    Dim strFontSize As String
    Dim strFontBold As String
    Dim strFontUnder As String
    Dim strFontItalic As String
    Dim strdata As String
    Dim strTitle As String
    Dim strPrtYN    As String
    Dim intPixeltoTwip As Long
    Dim intPixeltoTwipX As Long
    Dim intPixeltoTwipY As Long
    Dim varTmp As Variant
    
    If chkCorrect.Value = "1" Then
'        Call spdList.GetText(23, 1, varTmp): intPixeltoTwip = IIf(varTmp <> "", varTmp, 15)
'        Call spdList.GetText(23, 1, varTmp): intPixeltoTwipX = IIf(varTmp <> "", varTmp, 15)
'        Call spdList.GetText(24, 1, varTmp): intPixeltoTwipX = IIf(varTmp <> "", varTmp, 15)
    
        intPixeltoTwip = 14.405
        intPixeltoTwipX = 14.405
        intPixeltoTwipY = 14.405
    Else
        intPixeltoTwip = 15
        intPixeltoTwipX = 15
        intPixeltoTwipY = 15
    End If
    
    '-- 선택된 프린터로 출력
    For Each prtSelectPrinter In Printers
        If UCase(Trim(prtSelectPrinter.DeviceName)) = UCase(Trim(cmbPrinter.Text)) Then
            Set Printer = prtSelectPrinter
            boolPrinter_Select_Fales = True
            Exit For
        End If
    Next
    
    With spdList
        Printer.ScaleMode = vbTwips
        Picture1.AutoRedraw = True
        '-- 박스 그리기
        
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = 2
            Select Case Trim(.Text)
                Case "0"
                    Printer.ScaleMode = vbTwips
                    .Col = 4: strX1 = Trim(.Text) * intPixeltoTwip
                    .Col = 5: strX2 = Trim(.Text) * intPixeltoTwip
                    .Col = 6: strY1 = Trim(.Text) * intPixeltoTwip
                    .Col = 7: strY2 = Trim(.Text) * intPixeltoTwip

                    .Col = 8: strFont = Trim(.Text)
                    .Col = 9: strFontSize = Trim(.Text)
                    .Col = 10: strFontBold = Trim(.Text)
                    .Col = 11: strFontItalic = Trim(.Text)
                    .Col = 12: strFontUnder = Trim(.Text)
                    .Col = 22: strdata = Trim(.Text)


                    Printer.FontName = strFont
                    Printer.Font.Size = strFontSize
                    Printer.Font.Bold = IIf(strFontBold = "1", True, False)
                    Printer.Font.Italic = IIf(strFontItalic = "1", True, False)
                    Printer.Font.Underline = IIf(strFontUnder = "1", True, False)

                    Printer.CurrentX = strX1
                    Printer.CurrentY = strY1
                    Printer.Print strdata
                Case "1"
                    Printer.ScaleMode = vbTwips
                    .Col = 4: strX1 = Trim(.Text) * intPixeltoTwip
                    .Col = 5: strX2 = Trim(.Text) * intPixeltoTwip
                    .Col = 6: strY1 = Trim(.Text) * intPixeltoTwip
                    .Col = 7: strY2 = Trim(.Text) * intPixeltoTwip

                    .Col = 8: strFont = Trim(.Text)
                    .Col = 9: strFontSize = Trim(.Text)
                    .Col = 10: strFontBold = Trim(.Text)
                    .Col = 11: strFontItalic = Trim(.Text)
                    .Col = 12: strFontUnder = Trim(.Text)
                    .Col = 22: strdata = Trim(.Text)

                    Printer.FontName = strFont
                    Printer.Font.Size = strFontSize
                    Printer.Font.Bold = IIf(strFontBold = "1", True, False)
                    Printer.Font.Italic = IIf(strFontItalic = "1", True, False)
                    Printer.Font.Underline = IIf(strFontUnder = "1", True, False)

                    Printer.CurrentX = strX1
                    Printer.CurrentY = strY1
                    Printer.Print strdata

                Case "2"
                    Printer.ScaleMode = vbTwips
                    '.Col = 3: strTitle = Trim(.Text)
                    .Col = 29: strTitle = Trim(.Text)

                    .Col = 4: strX1 = Trim(.Text) * intPixeltoTwip
                    .Col = 5: strX2 = Trim(.Text) * intPixeltoTwip
                    .Col = 6: strY1 = Trim(.Text) * intPixeltoTwip
                    .Col = 7: strY2 = Trim(.Text) * intPixeltoTwip

'                    .Col = 8: strFont = Trim(.Text)
'                    .Col = 9: strFontSize = Trim(.Text)
'                    .Col = 17: strData = Trim(.Text)

                    Printer.PaintPicture Me.Controls(strTitle), strX1, strY1, strX2, strY2

                Case "3"
                    Printer.ScaleMode = vbTwips
                    
                    '.Col = 3: strTitle = Trim(.Text)
                    .Col = 29: strTitle = Trim(.Text)

                    .Col = 4: strX1 = Trim(.Text) * intPixeltoTwip
                    .Col = 5: strX2 = Trim(.Text) * intPixeltoTwip
                    .Col = 6: strY1 = Trim(.Text) * intPixeltoTwip
                    .Col = 7: strY2 = Trim(.Text) * intPixeltoTwip

'                    .Col = 8: strFont = Trim(.Text)
'                    .Col = 9: strFontSize = Trim(.Text)
'                    .Col = 17: strData = Trim(.Text)

                    Printer.PaintPicture Me.Controls(strTitle), strX1, strY1, strX2, strY2

                Case "4"
                    '.Col = 3: strTitle = Trim(.Text)
                    .Col = 29: strTitle = Trim(.Text)
                               strTitle = Mid(Trim(strTitle), 1, InStr(Trim(strTitle), "_IMG") - 1)


                    .Col = 4: strX1 = Trim(.Text) * intPixeltoTwip
                    .Col = 5: strX2 = Trim(.Text) * intPixeltoTwip
                    .Col = 6: strY1 = Trim(.Text) * intPixeltoTwip
                    .Col = 7: strY2 = Trim(.Text) * intPixeltoTwip

                    Dim x, y, W, H

                    Printer.ScaleMode = vbTwips
                    Printer.PSet (0, 0), vbWhite

                    x = Printer.ScaleX(strX1, vbTwips) ' X-position = 25 mm from left border
                    y = Printer.ScaleY(strY1, vbTwips)  ' Y-position = 25 mm from top border
                    W = Printer.ScaleX(strX2, vbTwips)  ' Width = 100 mm
                    H = Printer.ScaleY(strY2, vbTwips)  ' Height = 40 mm

                    '-- 바코드 회전
                    .Col = 16
                    Me.Controls(strTitle).Direction = IIf(Trim(.Text) = "0", 0, 1)
                    
                    Me.Controls(strTitle).PrinterScaleMode = vbTwips   '3:픽셀,1:트윕,6:밀리미터
                    Me.Controls(strTitle).Alignment = bcACenter
                    Me.Controls(strTitle).PrinterLeft = x '* 4.6
                    Me.Controls(strTitle).PrinterTop = y '* 5
                    Me.Controls(strTitle).PrinterWidth = W '(W * 5)  'W
                    Me.Controls(strTitle).PrinterHeight = H '(H * 5)  'H
                    Me.Controls(strTitle).PrinterHDC = Printer.hDC
                
                Case "5"
                    '-- 출력여부
                    .Col = 21: strPrtYN = Trim(.Text)
                    Printer.ScaleMode = vbTwips
                    
                    'If strPrtYN = "1" Then
                        
                        Printer.PSet (0, 0), vbWhite
                        
                        .Col = 4: strX1 = Trim(.Text) * intPixeltoTwip '* 13.3
                        .Col = 5: strX2 = Trim(.Text) * intPixeltoTwip '* 13.3
                        .Col = 6: strY1 = Trim(.Text) * intPixeltoTwip '* 13.3
                        .Col = 7: strY2 = Trim(.Text) * intPixeltoTwip '* 13.3
                        '선굵기
                        Printer.DrawWidth = 1
                        Printer.Line (strX1, strY1)-(strX2, strY2)
                    'End If
            End Select
        Next
    End With
    

    Printer.EndDoc
    
    'SavePicture Picture1.Image, "C:\TEST.BMP"
    
End Sub

Public Sub cmdSet_Click()

    '-- Mode Set [적용가능]
    If intMode = 1 Then
        Call objSet
    End If
    
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 동적 버튼 생성
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'''Private Sub Command1_Click()
'''
'''    Dim obj                 As Object
'''    Dim i                   As Integer
'''    Dim ClsEventObject      As ClassEventObject
'''
'''    ' 프로그램 정보 TextBox 숨김
'''    Text1.Visible = False
'''
'''    List1.Clear
'''
'''    ' 컬렉션 초기화
''''    Set m_ColCommandButton = Nothing
''''    Set m_ColCommandButton = New Collection
'''
'''    ' 동적 컨트롤 생성
'''    For i = 1 To Val(Combo1.Text)
'''        Set ClsEventObject = New ClassEventObject
'''
'''        If Option1.Value = True Then
'''            ' CommandButton
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectCommandButton, "DynamicCmd" & CStr(i))
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''            obj.Caption = "Button" & CStr(i)
'''        ElseIf Option2.Value = True Then
'''            ' TextBox
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectTextBox, "DynamicTxt" & CStr(i))
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''            obj.Text = "Text" & CStr(i)
'''        ElseIf Option3.Value = True Then
'''            ' Label
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLabel, "DynamicLbl" & CStr(i))
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''            obj.Caption = "Label" & CStr(i)
'''        ElseIf Option4.Value = True Then
'''            ' Image
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectImage, "DynamicImg" & CStr(i))
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''            obj.Picture = LoadPicture(App.Path & "\ugc.jpg")
'''
'''        ElseIf Option5.Value = True Then
'''            ' line
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectLine, "DynamicLine" & CStr(i))
'''            '-- 세로선
'''            obj.X1 = 100 * i
'''            obj.X2 = 100 * i
'''            obj.Y1 = 2070
'''            obj.Y2 = 4560
'''            '-- 가로선
'''            obj.X1 = 2850
'''            obj.X2 = 7080
'''            obj.Y1 = 100 * i
'''            obj.Y2 = 100 * i
'''
'''        Else
'''            ' barcode
'''            Set obj = ClsEventObject.CreateObject(Me, ClsEventMonitor, EventObjectBarcode, "DynamicBar" & CStr(i))
'''            obj.Alignment = bcACenter
'''            obj.Caption = "88006611"
'''            obj.Style = msSCode128B
'''            obj.Width = 3600
'''            obj.Height = 525
'''            obj.Top = 300 + (i - 1) * (525 + 30)
'''            obj.Left = 450
'''
''''            Barcod1.Alignment = bcACenter
'''            'Barcod1.Style = msSCode128B ' msS2of5
'''
'''        End If
'''
'''        obj.Visible = True
'''        'Set obj.Container = Frame2
'''        Set obj.Container = Picture1
'''
'''        m_ColCommandButton.Add ClsEventObject
'''
'''        Set ClsEventObject = Nothing
'''    Next
'''
'''End Sub


Private Sub MDIForm_Tool()
    
On Error GoTo ErrorRouten
    
    With tlbMain
        .AllowCustomize = False
        Set .ImageList = imlToolbar
        .TextAlignment = tbrTextAlignBottom '= tbrTextAlignRight
        .BorderStyle = ccNone
        .Appearance = cc3D
        .Style = tbrFlat
        Call .Buttons.Add(, TLBKEY_NEW, "", tbrDefault, "New")
        Call .Buttons.Add(, TLBKEY_OPEN, "", tbrDefault, "Open")
        Call .Buttons.Add(, TLBKEY_SAVE, "", tbrDefault, "Save")
        
        Call .Buttons.Add(, "", "", tbrSeparator)
        
        Call .Buttons.Add(, TLBKEY_MAKE, "", tbrDefault, "Make")
        Call .Buttons.Add(, TLBKEY_VIEW, "", tbrDefault, "View")
        Call .Buttons.Add(, "", "", tbrSeparator)
        Call .Buttons.Add(, TLBKEY_EDIT, "", tbrDefault, "Edit")
        Call .Buttons.Add(, TLBKEY_EXIT, "", tbrDefault, "Exit")
        Call .Buttons.Add(, "", "", tbrSeparator)
        
        
        .Refresh
    End With

Exit Sub

ErrorRouten:
'    Call ErrMsgProc(CallForm)

End Sub

Private Sub cmdUndo_Click()
    Dim Moveobj As Variant
    Dim x, y As Long
    
    Moveobj = LMousePos.obj
    x = LMousePos.fromx
    y = LMousePos.fromy

    Me.Controls(Moveobj).Left = x
    Me.Controls(Moveobj).Top = y

End Sub


Private Sub lblTitle_DblClick()
    
    If txtTag.Visible = True Then
        txtTag.Visible = False
    Else
        txtTag.Visible = True
    End If
    
End Sub

Private Sub mnuClose_Click()
        
    If MsgBox("종료하시겠습니까?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub mnuMake_Click()
    
    If MsgBox("작업파일을 생성하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
        Call MakeJOB
    End If
    
End Sub


' 첫번째 방법 : UTF-16을 나타내는 Byte Order Mark(BOM) 가 없을 경우,
'
Public Function UTF8FromUTF16(ByRef abytUTF16() As Byte) As Byte()
     
    Dim lngByteNum As Long
    Dim abytUTF8() As Byte
    Dim lngCharCount As Long
     
    On Error GoTo ConversionErr
     
    lngCharCount = (UBound(abytUTF16) + 1) \ 2
    ' UTF-16 LE 스트링의 문자의 수를 대입시켜, 변환에 필요한 바이트 수를 구합니다.
    lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, 0, 0, 0, 0)
                     
    If lngByteNum > 0 Then
        ' 변환된 코드를 반환받을 메모리를 확보한 후 함수를 호출합니다.
        ReDim abytUTF8(lngByteNum - 1)
        lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytUTF16(0), lngCharCount, _
                                         abytUTF8(0), lngByteNum, 0, 0)
        UTF8FromUTF16 = abytUTF8
    End If
    Exit Function
     
ConversionErr:
    MsgBox " Conversion failed "
    
End Function


' 두번째 방법 : BOM 을 무시한 후, UTF-8 방식으로 변환한 후,
'                    UTF-8 방식을 나타내는 Signature 를 추가하여 반환
'
Public Function UTF8FromUTF16withMark(ByRef abytUTF16() As Byte) As Byte()
    Dim abytTemp() As Byte
    Dim abytUTF8() As Byte
    Dim lngByteNum As Long
    Dim lngCharCount As Long
    Dim lngUpper As Long
     
    On Error GoTo ConversionErr
                   
    abytTemp = abytUTF16
    lngUpper = UBound(abytTemp)
    If lngUpper > 1 Then
        ' UTF-16 LE 의 바이트순서표식이 있을 경우 이를 일단 삭제합니다.
        ' &HFEFF 문자인데, LE에서는 도치되어 저장되므로, &HFF 가 먼저 위치함.
        If abytTemp(0) = &HFF And abytTemp(1) = &HFE Then
            Call CopyMemory(abytTemp(0), abytTemp(2), lngUpper - 1)
            ReDim Preserve abytTemp(lngUpper - 2)
            lngUpper = lngUpper - 2
        End If
    End If
    lngCharCount = (lngUpper + 1) \ 2

   ' 이제 변환에 필요한 메모리의 크기를 구합니다.
    lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytTemp(0), lngCharCount, 0, 0, 0, 0)
                     
    If lngByteNum > 0 Then
        ReDim abytUTF8(lngByteNum - 1)
        lngByteNum = WideCharToMultiByteArray(CP_UTF8, 0, abytTemp(0), lngCharCount, _
                                         abytUTF8(0), lngByteNum, 0, 0)
        lngUpper = UBound(abytUTF8)
        ' 변환되어 있는 UTF-8 바이트 배열 선두에 UTF-8 표식을 넣기 위해
        ' 기존의 바이트 배열을 뒤로 밀어내고, 배열 앞부분에 표식을 추가합니다.
        ReDim Preserve abytUTF8(lngUpper + 3)
        Call CopyMemory(abytUTF8(3), abytUTF8(0), lngUpper + 1)
        abytUTF8(0) = &HEF
        abytUTF8(1) = &HBB
        abytUTF8(2) = &HBF
         
        UTF8FromUTF16withMark = abytUTF8
    End If
    Exit Function
     
ConversionErr:
    MsgBox " Conversion failed "
    
End Function

Private Sub MakeLOF()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strdata As Variant
    Dim varTmp
    Dim abytUTF16() As Byte
    Dim abytUTF8() As Byte
    
    'Cancel을 True로 설정합니다.
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrHandler
    
    'Flags 속성을 설정합니다.
    CommonDialog1.flags = cdlCFEffects Or cdlCFBoth
    
    '[글꼴] 대화 상자를 표시합니다.
    CommonDialog1.ShowSave

    If Not LCase(Right(CommonDialog1.FileName, 4)) = ".lof" Then
        CommonDialog1.FileName = CommonDialog1.FileName & ".lof"
    End If
    
    Open CommonDialog1.FileName For Binary As #1
    With spdList
        strdata = ""
        For intRow = 1 To .MaxRows
            For intCol = 1 To .MaxCols - 1 '-- 마지막 Control제거
                .GetText intCol, intRow, varTmp: strdata = strdata & varTmp & "^"
            Next
            strdata = strdata & vbCr
        Next
        
    End With

    abytUTF16 = strdata
    'abytUTF16 = "유니코드 인코딩 변환 테스트 : UTF-16 LE 를 UTF-8 방식으로 변환하기"
    abytUTF8 = UTF8FromUTF16withMark(abytUTF16)
     
    'Open "C:\_UTF8TestFile.TXT" For Binary As #1
    Put #1, , abytUTF8
    Close #1
    'MsgBox " 변환 완료. " & vbCrLf & " 인터넷 익스플로러로 _UTF8TestFile.TXT 파일을 확인할 수 있습니다. "


    Close #1

    Exit Sub
    
ErrHandler:

End Sub

Private Sub MakeJOB()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strdata As Variant
    Dim varTmp
        
    On Error GoTo ErrHandler
    
    Open App.Path & "\" & gWork & "Job.txt" For Output As #1
        
    Print #1, "[JobPK]" & Chr(13) + Chr(10);
    Print #1, Me.Caption & ";" & Format(Now, "yyyy-mm-dd") & ";A;A;A;1;V" & Chr(13) + Chr(10);
    
    With spdList
        Print #1, "[S_Text]" & Chr(13) + Chr(10);
'        strData = ""
'        For intRow = 1 To .MaxRows
'            .GetText 2, intRow, varTmp
'            If varTmp = "0" Then
'                .GetText 3, intRow, varTmp
'                strData = strData & varTmp & ";"
'                .GetText 22, intRow, varTmp
'                strData = strData & varTmp
'                Print #1, strData & Chr(13) + Chr(10);
'                strData = ""
'            End If
'        Next
        
        '[D_Text]
        Print #1, "[D_Text]" & Chr(13) + Chr(10);
        strdata = ""
        For intRow = 1 To .MaxRows
            .GetText 2, intRow, varTmp
            If varTmp = "1" Then
                .GetText 3, intRow, varTmp
                strdata = strdata & varTmp & ";"
                .GetText 22, intRow, varTmp
                strdata = strdata & varTmp
                Print #1, strdata & Chr(13) + Chr(10);
                strdata = ""
            End If
        Next
        
        '[S_Image]
        Print #1, "[S_Image]" & Chr(13) + Chr(10);
        strdata = ""
        For intRow = 1 To .MaxRows
            .GetText 2, intRow, varTmp
            If varTmp = "2" Then
                .GetText 3, intRow, varTmp
                strdata = strdata & varTmp & ";"
                .GetText 17, intRow, varTmp
                'strData = strData & varTmp
                strdata = strdata & "0"
                Print #1, strdata & Chr(13) + Chr(10);
                strdata = ""
            End If
        Next
        
        '[D_Image]
        Print #1, "[D_Image]" & Chr(13) + Chr(10);
        strdata = ""
        For intRow = 1 To .MaxRows
            .GetText 2, intRow, varTmp
            If varTmp = "3" Then
                .GetText 3, intRow, varTmp
                strdata = strdata & varTmp & ";"
                .GetText 17, intRow, varTmp
                varTmp = Split(varTmp, "\")
                strdata = strdata & varTmp(UBound(varTmp))
                Print #1, strdata & Chr(13) + Chr(10);
                strdata = ""
            End If
        Next
        
        '[Barcode]
        Print #1, "[Barcode]" & Chr(13) + Chr(10);
        strdata = ""
        For intRow = 1 To .MaxRows
            .GetText 2, intRow, varTmp
            If varTmp = "4" Then
                .GetText 22, intRow, varTmp
                strdata = strdata & varTmp
                Print #1, strdata & Chr(13) + Chr(10);
                strdata = ""
            End If
        Next
        
    End With
    
    Close #1
    
    MsgBox Me.Caption & "의 작업파일이 생성되었습니다. ", vbOKOnly + vbInformation, Me.Caption

    Exit Sub
    
ErrHandler:

End Sub

''Private Sub mnuNew_Click()
''
''    Call FrmInitial
''
''    Dim sNo1, sNo2 As String
''    Dim intCnt As Integer
''    Dim strEditObjName As String
''    Dim strWLayout As String
''    Dim strHLayout As String
''
''AgainInput:
''
''    sNo1 = Mid(gLayOutValue(gLayOutUse), 1, InStr(gLayOutValue(gLayOutUse), ":") - 1) / 10
''    sNo2 = Mid(gLayOutValue(gLayOutUse), InStr(gLayOutValue(gLayOutUse), ":") + 1) / 10
''
'''    sNo1 = InputBox("라벨용지 높이를 입력하세요 [단위 : cm]", "높이 입력", "7.5")
'''
'''    If Len(sNo1) > 0 Then
'''        If Not IsNumeric(sNo1) Then
'''            MsgBox "숫자만 입력하세요.!", vbCritical
'''            GoTo AgainInput
'''        Else
'''            sNo2 = InputBox("라벨용지 넓이를 입력하세요 [단위 : cm]", "넓이 입력", "3.5")
'''            If Len(sNo2) > 0 Then
'''                If Not IsNumeric(sNo2) Then
'''                    MsgBox "숫자만 입력하세요.!", vbCritical
'''                    GoTo AgainInput
'''                End If
'''
'''            End If
'''        End If
'''    End If
''
''
''    If sNo1 <> "" And sNo2 <> "" Then
''        txtPaperHSize.Text = sNo1 '/ 10
''        txtPaperWSize.Text = sNo2 '/ 10
''
''        sNo1 = Round(sNo1 * CM_TOTWIP, 0)
''        sNo2 = Round(sNo2 * CM_TOTWIP, 0)
''
''        sstType.Tab = 5
''        '-- Left
''        txtTitle.Text = "LINE_L"    '항목명(뷰어)
''        txtTag.Text = "LINE_L"      '항목명(실제)
''        gblCtrlNm = "LINE_L"     '항목명(실제)
''        txtXpos.Text = "1"          'X 좌표
''        txtYpos.Text = "1"          'Y 좌표
''        txtLineHSize.Text = "1"     '선굵기
''        txtLineWSize.Text = sNo1   '라인폭
''        chkLineRotate.Value = "1"   '라인회전
''        chkPrint.Value = "0"        '출력여부
''
''        strEditObjName = objMake
''        If strEditObjName = "0" Then
''            '객체생성 성공
''            Call MakeSpdSaveList(txtTitle, sstType.Tab)
''        End If
''
''        '-- Right
''        txtTitle.Text = "LINE_R"    '항목명(뷰어)
''        txtTag.Text = "LINE_R"      '항목명(실제)
''        gblCtrlNm = "LINE_R"     '항목명(실제)
''        txtXpos.Text = sNo2          'X 좌표
''        txtYpos.Text = "1"          'Y 좌표
''        txtLineHSize.Text = "1"     '선굵기
''        txtLineWSize.Text = sNo1   '라인폭
''        chkLineRotate.Value = "1"   '라인회전
''        chkPrint.Value = "0"        '출력여부
''
''        strEditObjName = objMake
''        If strEditObjName = "0" Then
''            '객체생성 성공
''            Call MakeSpdSaveList(txtTitle, sstType.Tab)
''        End If
''
''        '-- Top
''        txtTitle.Text = "LINE_T"    '항목명(뷰어)
''        txtTag.Text = "LINE_T"      '항목명(실제)
''        gblCtrlNm = "LINE_T"     '항목명(실제)
''        txtXpos.Text = "1"          'X 좌표
''        txtYpos.Text = "1"          'Y 좌표
''        txtLineHSize.Text = "1"     '선굵기
''        txtLineWSize.Text = sNo2   '라인폭
''        chkLineRotate.Value = "0"   '라인회전
''        chkPrint.Value = "0"        '출력여부
''
''        strEditObjName = objMake
''        If strEditObjName = "0" Then
''            '객체생성 성공
''            Call MakeSpdSaveList(txtTitle, sstType.Tab)
''        End If
''
''        '-- Bottom
''        txtTitle.Text = "LINE_B"    '항목명(뷰어)
''        txtTag.Text = "LINE_B"      '항목명(실제)
''        gblCtrlNm = "LINE_B"     '항목명(실제)
''        txtXpos.Text = "1"          'X 좌표
''        txtYpos.Text = sNo1          'Y 좌표
''        txtLineHSize.Text = "1"     '선굵기
''        txtLineWSize.Text = sNo2   '라인폭
''        chkLineRotate.Value = "0"   '라인회전
''        chkPrint.Value = "0"        '출력여부
''
''        strEditObjName = objMake
''        If strEditObjName = "0" Then
''            '객체생성 성공
''            Call MakeSpdSaveList(txtTitle, sstType.Tab)
''        End If
''
''    End If
''
''End Sub

'Private Sub mnuSave_Click()
'    Dim i As Integer
'
'    Call MakeLOF
'
'End Sub

Private Sub mnuSet_Click()

    frmConfig.Show

End Sub

Private Sub mnuView_Click()

    'If MsgBox("작업파일을 생성하시겠습니까?", vbYesNo + vbInformation, Me.Caption) = vbYes Then
        Call MakeJOB
        
        Call Shell(App.Path & "\" & "NOTEPAD.EXE", vbNormalFocus)
        
        Me.WindowState = 1
        
    'End If

End Sub


Private Sub optDevide_Click(Index As Integer)
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim strBuf()    As String
    
    gDevide = optDevide(Index).Tag
    
    ' 컬렉션 초기화
    Set m_ColCommandButton = Nothing
    Set m_ColCommandButton = New Collection
    
    With spdList
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = 1
            Erase strBuf
            If Trim(.Text) <> "" Then
                ReDim Preserve strBuf(.MaxCols) As String
                For intCol = 2 To .MaxCols
                    .Col = intCol
                    strBuf(intCol - 1) = Trim(.Text)
                Next
                Call MakeLayout(strBuf)
                Erase strBuf
            End If
        Next
    End With
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'
'    If Button = 1 Then
'        Picture1.Cls '=============>다시 그리기
''        Picture1.CurrentX = X
''        Picture1.CurrentY = Y
'        DrawX = X '=========>눌려진좌표기억
'        DrawY = Y
'
'        Picture1.DrawMode = 10
'
'        Ot_X = X
'        Ot_Y = Y
'    End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Button = 1 Then
'        Picture1.DrawWidth = 1
'        Picture1.DrawStyle = 2
'
'        Picture1.Line (DrawX, DrawY)-(Ot_X, Ot_Y), vbBlack, B
'        Picture1.Line (DrawX, DrawY)-(X, Y), vbBlack, B
'
'        Ot_X = X
'        Ot_Y = Y
'    End If
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Button = 1 Then
'        Picture1.Line (DrawX, DrawY)-(Ot_X, Ot_Y), vbBlue, B
'        Picture1.DrawMode = 13
'        Picture1.DrawWidth = 1
'        Picture1.DrawStyle = 0 '========>단색(하지 않으면 그대로 점선)
'        Picture1.Line (DrawX, DrawY)-(X, Y), vbBlue, B
'    End If

End Sub

'-- 컨트롤 초기화
Private Sub CtrlInitial()
        
    txtPaperHSize.Text = ""
    txtPaperWSize.Text = ""
        
    '-- Tab 0
    txtFontName(0).Text = ""
    txtFontSize(0).Text = ""
    chkFontBold(0).Value = 0
    chkFontUnder(0).Value = 0
    chkFontItalic(0).Value = 0
    txtContent(0).Text = ""
    
    '-- Tab 1
    txtFontName(1).Text = ""
    txtFontSize(1).Text = ""
    chkFontBold(1).Value = 0
    chkFontUnder(1).Value = 0
    chkFontItalic(1).Value = 0
    txtContent(1).Text = ""
    
    '-- Tab 2
    txtImageName(0).Text = ""
    txtImageWSize(0).Text = ""
    txtImageHSize(0).Text = ""
    chkIStatic.Value = 0
    
    '-- Tab 3
    txtImageName(1).Text = ""
    txtImageWSize(1).Text = ""
    txtImageHSize(1).Text = ""
    
    '-- Tab 4
    txtBarDevide.Text = ""
    txtBarWSize.Text = ""
    txtBarHSize.Text = ""
    txtBarData.Text = ""
    chkBarRotate.Value = 0
    
    '-- Tab 5
    txtLineHSize.Text = ""
    txtLineWSize.Text = ""
    chkLineRotate.Value = 0
    
    gblCtrlNm = ""
    gblCtrlIdx = 0
    
    
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'

Private Function ShowOpen(Ufilter As String, Upath As String) As String
    
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Ufilter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = Upath
    OFName.lpstrTitle = "Open File"
    OFName.flags = 0

    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
        'ShowOpen = Mid(ShowOpen, 1, Len(ShowOpen) - 1)
    Else
        ShowOpen = ""
    End If
    
End Function


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

'-- 오픈한 LOF 파일을 스프레드에 표시한다,
'-- 용도 : 적용,저장시 사용한다.
Private Sub SetList(varBuf As Variant)
    Dim intCnt As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    
    With spdList
        .MaxRows = .MaxRows + 1
        intRow = .MaxRows
        For intCnt = 0 To UBound(varBuf) '- 1
            If .MaxRows = 1 And intCnt = 0 Then
                If Len(varBuf(intCnt)) > 1 Then varBuf(intCnt) = Right(varBuf(intCnt), 1)
                .SetText intCnt + 1, intRow, CStr(varBuf(intCnt))
            Else
                If intCnt = UBound(varBuf) Then
                    If varBuf(1) = "4" Then
                        .SetText intCnt + 1, intRow, strBarImgName
                    Else
                        .SetText intCnt + 1, intRow, Trim(txtTag.Text)
                    End If
                Else
                    .SetText intCnt + 1, intRow, CStr(varBuf(intCnt))
                End If
            End If
        Next
    End With

End Sub

Private Function BarIdxMapper(idx As Variant) As String
    

    Select Case idx
    Case 0:     BarIdxMapper = 3
    Case 1:     BarIdxMapper = 5
    Case 2:     BarIdxMapper = ""
    Case 3:     BarIdxMapper = 11
    Case 4:     BarIdxMapper = 12
    Case 5:     BarIdxMapper = 1
    Case 6:     BarIdxMapper = 2
    Case 7:     BarIdxMapper = 10
    Case 8:     BarIdxMapper = 22
    Case 9:     BarIdxMapper = 4
    Case 10:    BarIdxMapper = 18
    Case 11:    BarIdxMapper = 6
    Case 12:    BarIdxMapper = 7
    Case 13:    BarIdxMapper = 8
    Case 14:    BarIdxMapper = ""
    Case 15:    BarIdxMapper = 9
    Case 16:    BarIdxMapper = 20
    Case 17:    BarIdxMapper = 13
    Case 18:    BarIdxMapper = 14
    Case 19:    BarIdxMapper = ""
    Case 20:    BarIdxMapper = ""
    Case 21:    BarIdxMapper = ""
    Case 22:    BarIdxMapper = ""
    Case 23:    BarIdxMapper = 15
    Case 24:    BarIdxMapper = ""
    Case 25:    BarIdxMapper = ""
    Case 26:    BarIdxMapper = ""
    Case 27:    BarIdxMapper = ""
    Case 28:    BarIdxMapper = ""
    Case Else:  BarIdxMapper = ""
    End Select



End Function

'-- 구분별로 오프젝트 내역을 각 항목에 표시한다.
'   구분[varBuf(1)] 0:SText,1:DText,2:SImage,3:DImage,4:Barcode,5:Line
Private Sub MakeLayout(varBuf As Variant)
    Dim strEditObjName      As String
    Dim i As Integer
    Dim strFVar As String
    Dim strTmp
    
MakeAgain:
    
    sstType.Tab = varBuf(1)
    
    txtPaperHSize.Text = varBuf(25)
    txtPaperWSize.Text = varBuf(25)
    
    strFVar = ""
    For i = 1 To Len(varBuf(0))
        If Asc(Mid(varBuf(0), i, 1)) <> 63 Then
           strFVar = strFVar & Mid(varBuf(0), i, 1)
        Else
            'Stop
        End If
    Next
    
    Select Case varBuf(1)
        Case 0  '## Static Label ##
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '항목명(뷰어)
            txtXpos.Text = varBuf(3)                            'X 좌표
            txtYpos.Text = varBuf(5)                            'Y 좌표
            txtFontName(0).Text = varBuf(7)                     '폰트명
            txtFontSize(0).Text = varBuf(8)                     '폰트크기
            chkFontBold(0).Value = varBuf(9)                    '    굵게
            chkFontUnder(0).Value = varBuf(11)                  '    밑줄
            chkFontItalic(0).Value = varBuf(10)                 '    기울게
            txtContent(0).Text = varBuf(21)                     'Text
            chkTStatic.Value = varBuf(26)                       '무조건고정
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '출력안함
        
        Case 1  '## Dynamic Label ##
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '항목명(뷰어)
            txtXpos.Text = varBuf(3)                            'X 좌표
            txtYpos.Text = varBuf(5)                            'Y 좌표
            txtFontName(1).Text = varBuf(7)                     '폰트명
            txtFontSize(1).Text = varBuf(8)                     '폰트크기
            chkFontBold(1).Value = varBuf(9)                    '    굵게
            chkFontUnder(1).Value = varBuf(11)                  '    밑줄
            chkFontItalic(1).Value = varBuf(10)                 '    기울게
            txtContent(1).Text = varBuf(21)                     'Text
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '출력안함
        
        Case 2  '## Static Image ##
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '항목명(뷰어)
            txtXpos.Text = varBuf(3)                            'X 좌표
            txtYpos.Text = varBuf(5)                            'Y 좌표
            txtImageName(0).Text = varBuf(16)                   '이미지경로
            txtImageWSize(0).Text = varBuf(4)                   '      가로SIZE
            txtImageHSize(0).Text = varBuf(6)                   '      세로SIZE
            chkIStatic.Value = varBuf(26)                       '무조건고정
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '출력안함
                        
        Case 3  '## Dynamic Image ##
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '항목명(뷰어)
            txtXpos.Text = varBuf(3)                            'X 좌표
            txtYpos.Text = varBuf(5)                            'Y 좌표
            txtImageName(1).Text = varBuf(16)                   '이미지경로
            txtImageWSize(1).Text = varBuf(4)                   '      가로SIZE
            txtImageHSize(1).Text = varBuf(6)                   '      세로SIZE
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '출력안함
            
        Case 4
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '항목명(뷰어)
            
            
            '-- 바코드 타입 기존 프로그램과 신규프로그램 Mapping
            strTmp = BarIdxMapper(varBuf(13))
            If strTmp = "" Then
                cboBarType.ListIndex = 7                   '바코드 타입
            Else
                cboBarType.ListIndex = strTmp                   '바코드 타입
            End If
            
            txtXpos.Text = varBuf(3)                            'X 좌표
            txtYpos.Text = varBuf(5)                            'Y 좌표
            txtBarData.Text = varBuf(21)                        '바코드Data
            txtBarWSize.Text = varBuf(4)                        '      길이SIZE
            txtBarHSize.Text = varBuf(6)                        '      세로SIZE
            chkBarRotate.Value = IIf(varBuf(15) = "0", "0", "1") '     회전
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '출력안함
        
        Case 5
            'txtTag.Text = Replace(varBuf(2), "-", "_")          '항목명(실제)
            txtTag.Text = "Control_" & strFVar
            txtTitle.Text = varBuf(2)                           '항목명(뷰어)
            txtXpos.Text = varBuf(3)                            'X 좌표
            txtYpos.Text = varBuf(5)                            'Y 좌표
            chkLineRotate.Value = IIf(varBuf(17) = "0", "0", "1")   '라인회전
            txtLineHSize.Text = varBuf(18)                      '선굵기
            txtLineWSize.Text = varBuf(19)                      '라인폭
            chkPrint.Value = IIf(varBuf(20) = "1", "0", "1")    '출력안함
    End Select
    
    '-- 객체이름 업데이트
    gblCtrlNm = txtTag.Text
    gblCtrlIdx = strFVar
    
    '-- 객체생성
    strEditObjName = objMake
    
    If strEditObjName = "0" Then
        '객체생성 성공
    Else
        '객체생성 실패
        varBuf(2) = strEditObjName
        GoTo MakeAgain
    End If

End Sub


Private Sub SetLayout(intTabidx As Integer)

    '구분[varBuf(1)] 0:SText,1:DText,2:SImage,3:DImage,4:Barcode,5:Line
    
    Dim intCnt As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strIdx As String
    Dim strTitle As String
    
    With spdList
        For intRow = 1 To .MaxRows
            '항목구분,항목명 비교
            .Row = intRow
            .Col = 2: strIdx = Trim(.Text)
            .Col = 29: strTitle = Trim(.Text)
            If intTabidx = strIdx And Trim(txtTag.Text) = Trim(strTitle) Then
                Select Case intTabidx
                    Case 0
                        .SetText 4, intRow, txtXpos.Text
                        .SetText 6, intRow, txtYpos.Text
                        .SetText 8, intRow, txtFontName(0).Text
                        .SetText 9, intRow, txtFontSize(0).Text
                        .SetText 10, intRow, IIf(chkFontBold(0).Value = "0", "0", "1")
                        .SetText 11, intRow, IIf(chkFontItalic(0).Value = "0", "0", "1")
                        .SetText 12, intRow, IIf(chkFontUnder(0).Value = "0", "0", "1")
                        .SetText 22, intRow, Trim(txtContent(0).Text)
                        .SetText 21, intRow, IIf(chkPrint.Value = "1", "0", "1")      '출력여부
                        .SetText 27, intRow, IIf(chkTStatic.Value = "0", "0", "1")      '무조건고정
            
                    Case 1
                        .SetText 4, intRow, txtXpos.Text
                        .SetText 6, intRow, txtYpos.Text
                        .SetText 8, intRow, txtFontName(1).Text
                        .SetText 9, intRow, txtFontSize(1).Text
                        .SetText 10, intRow, IIf(chkFontBold(1).Value = "0", "0", "1")
                        .SetText 11, intRow, IIf(chkFontItalic(1).Value = "0", "0", "1")
                        .SetText 12, intRow, IIf(chkFontUnder(1).Value = "0", "0", "1")
                        .SetText 22, intRow, Trim(txtContent(1).Text)
                        .SetText 21, intRow, IIf(chkPrint.Value = "1", "0", "1")      '출력여부
            
                    Case 2
                        .SetText 4, intRow, txtXpos.Text
                        .SetText 5, intRow, txtImageWSize(0).Text
                        .SetText 6, intRow, txtYpos.Text
                        .SetText 7, intRow, txtImageHSize(0).Text
                        .SetText 17, intRow, txtImageName(0).Text
                        
                        .SetText 21, intRow, IIf(chkPrint.Value = "1", "0", "1")      '출력여부
                        .SetText 27, intRow, IIf(chkIStatic.Value = "0", "0", "1")      '무조건고정
            
                    Case 3
                        .SetText 4, intRow, txtXpos.Text
                        .SetText 5, intRow, txtImageWSize(1).Text
                        .SetText 6, intRow, txtYpos.Text
                        .SetText 7, intRow, txtImageHSize(1).Text
                        .SetText 17, intRow, txtImageName(1).Text
                        
                        .SetText 21, intRow, IIf(chkPrint.Value = "1", "0", "1")      '출력여부
            
                    Case 4
                        .SetText 4, intRow, txtXpos.Text
                        .SetText 5, intRow, txtBarWSize.Text
                        .SetText 6, intRow, txtYpos.Text
                        .SetText 7, intRow, txtBarHSize.Text
                        .SetText 14, intRow, cboBarType.ListIndex    '-- 바코드 종류
                        '.SetText 15, intRow, cboBarType.ListIndex    '-- 바코드 폭
                        .SetText 16, intRow, IIf(chkBarRotate.Value = "0", "0", "2")     '-- 바코드 회전
                        .SetText 22, intRow, Trim(txtBarData.Text)     '-- 바코드 출력값
                        
                        .SetText 21, intRow, IIf(chkPrint.Value = "1", "0", "1")        '출력여부
                    
                    Case 5
                        .SetText 4, intRow, txtXpos.Text
                        .SetText 5, intRow, txtXpos.Text
                        .SetText 6, intRow, txtYpos.Text
                        .SetText 7, intRow, txtLineWSize.Text
                        .SetText 9, intRow, txtLineHSize.Text
                        .SetText 18, intRow, IIf(chkLineRotate.Value = "0", "0", "1")   '라인회전
                        .SetText 19, intRow, txtLineHSize.Text                          '라인두께
                        .SetText 20, intRow, txtLineWSize.Text                          '라인폭
    
                        .SetText 21, intRow, IIf(chkPrint.Value = "1", "0", "1")        '출력여부
            
                End Select
                
                Exit Sub
            End If
        Next
    End With
    
    


End Sub


Public Function toUTF8(ByVal szSource As String) As String
On Error GoTo ErrHandler

Dim szChar As String
Dim WideChar As Long
Dim nLength As Integer
Dim i As Integer

    nLength = Len(szSource)
    
    For i = 1 To nLength
        szChar = Mid(szSource, i, 1)
        
        If Asc(szChar) < 0 Then
            WideChar = CLng(AscB(MidB(szChar, 2, 1))) * 256 + AscB(MidB(szChar, 1, 1))
        
            If (WideChar And &HFF80) = 0 Then
                toUTF8 = toUTF8 & Hex(WideChar)
            ElseIf (WideChar And &HF000) = 0 Then
                toUTF8 = toUTF8 & _
                Hex(CInt((WideChar And &HFFC0) / 64) Or &HC0) & _
                Hex(WideChar And &H3F Or &H80)
            Else
                toUTF8 = toUTF8 & _
                Hex(CInt((WideChar And &HF000) / 4096) Or &HE0) & _
                Hex(CInt((WideChar And &HFFC0) / 64) And &H3F Or &H80) & _
                Hex(WideChar And &H3F Or &H80)
        
            End If
        Else
            toUTF8 = toUTF8 & Hex(Asc(szChar))
        End If
    Next

Exit Function

ErrHandler:
    toUTF8 = ""

End Function

Public Function URLEncode(URLStr As String) As String

Dim sURL        As String   '** 입력받은 URL 문자열
Dim sBuffer     As String   '** URL 인코딩 처리 중 URL 을 담을 버퍼 문자열
Dim sTemp       As String   '** 임시 문자열
Dim cChar       As String   '** URL 문자열 중 현재 인텍스의 문자
Dim lErrNum     As Long     '** 오류 번호
Dim sErrSource  As String   '** 오류 소스
Dim sErrDesc    As String   '** 소류 설명
Dim sMsg        As String   '** 오류 메세지
Dim Index       As Integer

On Error GoTo ErrorHanddle:

    sURL = Trim(URLStr) '** URL 문자열을 얻는다.
    sBuffer = "" '** 임시 버퍼용 문자열 변수 초기화.

    '******************************************************
    '* URL 인코딩 작업
    '******************************************************

    For Index = 1 To Len(sURL)
        '** 현재 인덱스의 문자를 얻는다.
        cChar = Mid(sURL, Index, 1)
        
        If cChar = "0" Or (cChar >= "1" And cChar <= "9") Or (cChar >= "a" And cChar <= "z") Or (cChar >= "A" And cChar <= "Z") Or _
                          cChar = "-" Or cChar = "_" Or cChar = "." Or cChar = "*" Then
            '** URL 에 허용되는 문자들 :: 버퍼 문자열에 추가한다.
            sBuffer = sBuffer & cChar
        ElseIf cChar = " " Then
            '** 공백 문자 :: + 로 대체하여 버퍼 문자열에 추가한다.
            sBuffer = sBuffer & "+"
        Else
            '** URL 에 허용되지 않는 문자들 :: % 로 인코딩해서 버퍼 문자열에 추가한다.
            sTemp = CStr(Hex(Asc(cChar)))
            If Len(sTemp) = 4 Then
                sBuffer = sBuffer & "%" & Left(sTemp, 2) & "%" & Mid(sTemp, 3, 2)
            ElseIf Len(sTemp) = 2 Then
                sBuffer = sBuffer & "%" & sTemp
            End If
        End If
    Next

    '** 결과를 리턴한다.
    URLEncode = sBuffer

Exit Function

ErrorHanddle:

    '** 오류가 발생하면 공백 문자를 리턴한다.
    URLEncode = ""
    
    '** 오류 정보를 얻는다.
    lErrNum = Err.Number
    sErrSource = Err.Source
    sErrDesc = Err.Description
    
    '** 이벤트 로그에 오류를 기록한다.
    sMsg = vbCrLf & vbCrLf & _
    "Error Object : EgoCube.URLTools," & vbCrLf & _
    "Error Method : Public Function URLEncode(URLStr As String) As String," & vbCrLf & _
    "Error Number : " & lErrNum & "," & vbCrLf & _
    "Error Source : " & sErrSource & "," & vbCrLf & _
    "Error Description : " & sErrDesc
    
    App.LogEvent sMsg, vbLogEventTypeError
    
    '** 오류를 발생시킨다.
    Err.Raise lErrNum, sErrSource, sErrDesc
    

Exit Function


End Function

'Private Sub mnuOpen_Click()
'    Dim strSrcfile  As Variant
'    Dim varBuffer() As Variant
'    Dim varBuf      As Variant
'    Dim lngBufLen   As Long
'    Dim i           As Long
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim j           As Long
'    Dim bytBuff()   As Byte
'
'    Static ChkSumCnt As Long
'    Dim strTxt As String
'
'    Dim FileNumber As Long
'    Dim FileName As String
'    Dim FileCount As Long
'    Dim LineCount As Long
'    Dim FileOpenNumber As Integer
'    Dim data As String
'    Dim splitdata() As String
'
'    Dim utf8() As Byte
'    Dim ucs2 As Variant
'    Dim chars As Long
'    Dim varTmp As Variant
'
'    ' 폼초기화
'    Call FrmInitial
'
'    'Cancel을 True로 설정합니다.
'    CommonDialog1.CancelError = True
'    On Error GoTo ErrHandler
'
'    '경로 속성을 설정합니다.
'    CommonDialog1.InitDir = App.Path & "\" & gLayOut
'    CommonDialog1.Filter = "LayoutFile(*.lof)|*.lof"
'
'    '[파일] 대화 상자를 표시합니다.
'    CommonDialog1.ShowOpen
'    strSrcfile = CommonDialog1.FileName
'
'    '컬렉션 초기화
'    Set m_ColCommandButton = Nothing
'    Set m_ColCommandButton = New Collection
'
'    'LOF 파일 열기
'    FileName = CommonDialog1.FileName
'    varTmp = Split(FileName, "\")
'    Me.Caption = varTmp(UBound(varTmp))
'    FileOpenNumber = FreeFile()
'    LineCount = 0
'
'    Open FileName For Binary As #1   'UTF-8 문서지정
'    ReDim utf8(LOF(1))
'
'    Get #1, , utf8
'
'    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), 0, 0)
'    ucs2 = Space(chars)
'    chars = MultiByteToWideChar(CP_UTF8, 0, VarPtr(utf8(0)), LOF(1), StrPtr(ucs2), chars)
'    varBuf = Split(ucs2, Chr(13))
'
'    Close #1
'
'
'    '오픈한 LOF파일 버퍼에 쓰기
'    For i = 0 To UBound(varBuf)
'        ReDim Preserve varBuffer(i)
'        varBuffer(LineCount) = varBuf(i)
'        LineCount = LineCount + 1
'    Next
'
'    '오픈한 LOF파일 화면그리기/스프레드쓰기
'    For i = 0 To UBound(varBuffer) - 1
'        If varBuffer(i) <> "" Then
'            varBuf = Split(varBuffer(i), "^")
'            Call MakeLayout(varBuf)
'            Call SetList(varBuf)
'        End If
'    Next
'
''    intMode = 1
'
'    Exit Sub
'
'ErrHandler:
'
'End Sub



Private Sub spdList_Click(ByVal Col As Long, ByVal Row As Long)
        
    Call SetControl(Row)
    
End Sub

Private Sub SetControl(intRow As Long)

Dim strTmp As String

    With spdList
        .Row = intRow
        '-- 제목
        .Col = 2:   sstType.Tab = Trim(.Text)
        .Col = 3:   txtTitle.Text = Trim(.Text)
        .Col = 29:  txtTag.Text = Trim(.Text)
        '-- 위치
        .Col = 4:   txtXpos.Text = Trim(.Text)
        .Col = 6:   txtYpos.Text = Trim(.Text)
        '-- 넓이,높이(두께)
        Select Case sstType.Tab
            Case 2: .Col = 5:  txtImageWSize(0).Text = Trim(.Text)
                    .Col = 7:  txtImageHSize(0).Text = Trim(.Text)
            Case 3: .Col = 5:  txtImageWSize(1).Text = Trim(.Text)
                    .Col = 7:  txtImageHSize(1).Text = Trim(.Text)
            Case 4: .Col = 5:  txtBarWSize.Text = Trim(.Text)
                    .Col = 7:  txtBarHSize.Text = Trim(.Text)
        End Select
        '-- 폰트
        Select Case sstType.Tab
            Case 0: .Col = 8:  txtFontName(0).Text = Trim(.Text)
                    .Col = 9:  txtFontSize(0).Text = Trim(.Text)
                    .Col = 10: chkFontBold(0).Value = IIf(Trim(.Text) = "0", "0", "1")   '폰트굵게
                    .Col = 11: chkFontUnder(0).Value = IIf(Trim(.Text) = "0", "0", "1")  '폰트밑줄
                    .Col = 12: chkFontItalic(0).Value = IIf(Trim(.Text) = "0", "0", "1") '폰트기울게
                    '.Col = 13: chkFontItalic(0).Value = IIf(Trim(.Text) = "0", "0", "1") '폰트회전
            Case 1: .Col = 8:  txtFontName(1).Text = Trim(.Text)
                    .Col = 9:  txtFontSize(1).Text = Trim(.Text)
                    .Col = 10: chkFontBold(1).Value = IIf(Trim(.Text) = "0", "0", "1")   '폰트굵게
                    .Col = 11: chkFontUnder(1).Value = IIf(Trim(.Text) = "0", "0", "1")  '폰트밑줄
                    .Col = 12: chkFontItalic(1).Value = IIf(Trim(.Text) = "0", "0", "1") '폰트기울게
                    '.Col = 13: chkFontItalic(0).Value = IIf(Trim(.Text) = "0", "0", "1") '폰트회전
        End Select
        '-- 바코드
        '-- 바코드 타입 기존 프로그램과 신규프로그램 Mapping
        .Col = 14:   strTmp = BarIdxMapper(Trim(.Text))
        If strTmp = "" Then
            cboBarType.ListIndex = 7
        Else
            cboBarType.ListIndex = strTmp
        End If
        .Col = 15:  txtBarDevide.Text = Trim(.Text)
        .Col = 16:  chkBarRotate.Value = IIf(Trim(.Text) = "0", 0, 2)
        '-- 이미지
        If sstType.Tab = 3 Then
            .Col = 17:  txtImageName(0).Text = Trim(.Text)
        ElseIf sstType.Tab = 4 Then
            .Col = 17:  txtImageName(1).Text = Trim(.Text)
        End If
        '-- 라인
        .Col = 18:  chkLineRotate.Value = IIf(Trim(.Text) = "0", 0, 1)
        .Col = 19:  txtLineHSize.Text = Trim(.Text)
        .Col = 20:  txtLineWSize.Text = Trim(.Text)
        '-- 출력여부
        .Col = 21:  chkPrint.Value = IIf(Trim(.Text) = "1", 0, 1)
        '-- 출력값
        Select Case sstType.Tab
            Case 0:     .Col = 22: txtContent(0).Text = Trim(.Text)
            Case 1:     .Col = 22: txtContent(1).Text = Trim(.Text)
            Case 4:     .Col = 22: txtBarData.Text = Trim(.Text)
        End Select
        '-- 무조건고정
        If sstType.Tab = 0 Then
            .Col = 27:  chkTStatic.Value = IIf(Trim(.Text) = "0", 0, 1)
        ElseIf sstType.Tab = 2 Then
            .Col = 27:  chkIStatic.Value = IIf(Trim(.Text) = "0", 0, 1)
        End If
        
    End With

End Sub


Private Sub spdList_KeyPress(KeyAscii As Integer)
    Dim varTmp As Variant
        
    If KeyAscii = 13 Then
        
        Call SetControl(spdList.ActiveRow)
        
        intMode = 1
        
        Call cmdSet_Click
    
    End If

End Sub

'Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If X > (Command2.Width - 100) And Y > (Command2.Height - 100) And Button = vbLeftButton Then
'        drageMode = True
'    Else
'        drageMode = False
'    End If
'    If drageMode Then
'        Command2.Height = Y
'        Command2.Width = X
'    End If
'End Sub


Private Sub sstType_Click(PreviousTab As Integer)
    Select Case sstType.Tab
        Case 0
            txtTitle.Text = "S_TEXT" & gblCtrlIdx
            'cmdFont(0).SetFocus
        Case 1
            txtTitle.Text = "D_TEXT" & gblCtrlIdx
            'cmdFont(1).SetFocus
        Case 2
            txtTitle.Text = "S_Image" & gblCtrlIdx
            'cmdImage(0).SetFocus
        Case 3
            txtTitle.Text = "D_Image" & gblCtrlIdx
            'cmdImage(1).SetFocus
        Case 4
            txtTitle.Text = "BARCODE" & gblCtrlIdx
            'cboBarType.SetFocus
        Case 5
            txtTitle.Text = "LINE" & gblCtrlIdx
            'txtLineHSize.SetFocus
            txtLineHSize.Text = "1"
    End Select
    
    txtTag.Text = ""
    txtXpos.Text = 10
    txtYpos.Text = 10
    
    cboType.ListIndex = sstType.Tab

End Sub



Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case TLBKEY_NEW
            Call mnuNew_Click
        Case TLBKEY_OPEN
            Call mnuOpen_Click
        Case TLBKEY_SAVE
            Call mnuSave_Click
        Case TLBKEY_MAKE
            Call mnuMake_Click
        Case TLBKEY_VIEW
            Call mnuView_Click
        Case TLBKEY_EDIT
            Call mnuSet_Click
        Case TLBKEY_EDIT
            Call mnuSet_Click
        Case TLBKEY_EXIT
            Call mnuClose_Click
    End Select

End Sub

Private Sub tmrMove_Timer()
    
    Call objMove(intMoveIdx)

End Sub


Private Sub txtBarHSize_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtBarHSize.Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtBarHSize.SetFocus
        End If
    End If
    
End Sub

Private Sub txtBarWSize_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtBarWSize.Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtBarWSize.SetFocus
        End If
    End If

End Sub

Private Sub txtDevide_KeyPress(KeyAscii As Integer)
    Dim intRow      As Integer
    Dim intCol      As Integer
    Dim strBuf()    As String
    
    If KeyAscii = 13 Then
        If IsNumeric(txtDevide.Text) Then
            gDevide = txtDevide.Text
            
            ' 컬렉션 초기화
            Set m_ColCommandButton = Nothing
            Set m_ColCommandButton = New Collection
            
            With spdList
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = 1
                    Erase strBuf
                    If Trim(.Text) <> "" Then
                        ReDim Preserve strBuf(.MaxCols) As String
                        For intCol = 2 To .MaxCols
                            .Col = intCol
                            strBuf(intCol - 1) = Trim(.Text)
                        Next
                        Call MakeLayout(strBuf)
                        Erase strBuf
                    End If
                Next
            End With
        Else
            MsgBox "숫자만 입력이 가능합니다.", vbInformation, Me.Caption
            txtDevide.SetFocus
            Exit Sub
        End If
    End If
End Sub


Private Sub txtFontSize_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtFontSize(Index).Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtFontSize(Index).SetFocus
        End If
    End If

End Sub


Private Sub txtImageDevide_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Call cmdImageDevSet_Click(Index)
    End If
    
End Sub

Private Sub txtImageHSize_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtImageHSize(Index).Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtImageHSize(Index).SetFocus
        End If
    End If

End Sub

Private Sub txtImageWSize_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtImageWSize(Index).Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtImageWSize(Index).SetFocus
        End If
    End If

End Sub

Private Sub txtLineHSize_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtLineHSize.Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtLineHSize.SetFocus
        End If
    End If

End Sub

Private Sub txtLineWSize_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtLineWSize.Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtLineWSize.SetFocus
        End If
    End If

End Sub

Private Sub txtPaperHSize_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtPaperHSize.Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtPaperHSize.SetFocus
        End If
    End If

End Sub

Private Sub txtPaperWSize_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtPaperWSize.Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtPaperWSize.SetFocus
        End If
    End If

End Sub

Private Sub txtXpos_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtXpos.Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtXpos.SetFocus
        End If
    End If

End Sub

Private Sub txtYpos_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Not IsNumeric(Trim(txtYpos.Text)) Then
            MsgBox "숫자만 입력이 가능합니다.", vbOKOnly + vbInformation, Me.Caption
            txtYpos.SetFocus
        End If
    End If

End Sub



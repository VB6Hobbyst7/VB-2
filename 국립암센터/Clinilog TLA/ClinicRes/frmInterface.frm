VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInterface 
   Caption         =   "CLINILOG Interface Program [Service Center ☎(02)6205-1751]"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15240
   Icon            =   "frmInterface.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   15240
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8910
      TabIndex        =   64
      Top             =   1260
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtBuff 
      Height          =   345
      Left            =   1050
      TabIndex        =   5
      Text            =   $"frmInterface.frx":0442
      Top             =   1260
      Visible         =   0   'False
      Width           =   7545
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   8565
      Left            =   2820
      TabIndex        =   63
      Top             =   1830
      Visible         =   0   'False
      Width           =   10215
      _Version        =   393216
      _ExtentX        =   18018
      _ExtentY        =   15108
      _StockProps     =   64
      ColHeaderDisplay=   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   100
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":0545
   End
   Begin VB.CheckBox chkAll 
      Height          =   255
      Left            =   750
      TabIndex        =   61
      Top             =   960
      Width           =   195
   End
   Begin FPSpread.vaSpread vasPrint_1 
      Height          =   8565
      Left            =   3060
      TabIndex        =   44
      Top             =   1230
      Visible         =   0   'False
      Width           =   10215
      _Version        =   393216
      _ExtentX        =   18018
      _ExtentY        =   15108
      _StockProps     =   64
      ColHeaderDisplay=   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   10
      MaxRows         =   100
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":8948
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8850
      TabIndex        =   38
      Top             =   9840
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Height          =   6435
      Left            =   1980
      TabIndex        =   10
      Top             =   2970
      Visible         =   0   'False
      Width           =   10995
      Begin VB.TextBox txtEquip 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2520
         Width           =   1545
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "◀"
         Height          =   465
         Left            =   810
         TabIndex        =   27
         Top             =   4380
         Width           =   645
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "▶"
         Height          =   465
         Left            =   1470
         TabIndex        =   26
         Top             =   4380
         Width           =   645
      End
      Begin VB.TextBox txtRack 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2970
         Width           =   1545
      End
      Begin VB.TextBox txtTube 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3390
         Width           =   1545
      End
      Begin VB.TextBox txtResDate 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2070
         Width           =   2475
      End
      Begin VB.TextBox txtPName 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1275
         Width           =   1545
      End
      Begin VB.TextBox txtPID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   825
         Width           =   1545
      End
      Begin VB.TextBox txtID 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   390
         Width           =   1545
      End
      Begin FPSpread.vaSpread vasRes1 
         Height          =   5865
         Left            =   2940
         TabIndex        =   11
         Top             =   330
         Width           =   3885
         _Version        =   393216
         _ExtentX        =   6853
         _ExtentY        =   10345
         _StockProps     =   64
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   20
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "frmInterface.frx":10DE2
      End
      Begin FPSpread.vaSpread vasRes2 
         Height          =   5865
         Left            =   6870
         TabIndex        =   12
         Top             =   330
         Width           =   3885
         _Version        =   393216
         _ExtentX        =   6853
         _ExtentY        =   10345
         _StockProps     =   64
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   20
         RowHeaderDisplay=   0
         ScrollBars      =   0
         SpreadDesigner  =   "frmInterface.frx":11405
      End
      Begin Threed.SSCommand cmdCloseDetail 
         Height          =   495
         Left            =   1560
         TabIndex        =   28
         Top             =   5175
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   873
         _StockProps     =   78
         Caption         =   "닫기"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "W/L No"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   2580
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Rack"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   30
         Top             =   3030
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Tube"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   29
         Top             =   3450
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과시간"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   19
         Top             =   1785
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자이름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   17
         Top             =   1335
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "등록번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   15
         Top             =   885
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   450
         Width           =   840
      End
   End
   Begin VB.Frame frameSch 
      Height          =   9525
      Left            =   30
      TabIndex        =   34
      Top             =   1770
      Visible         =   0   'False
      Width           =   15075
      Begin VB.CommandButton cmdSchClose 
         Caption         =   "닫기"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   13440
         TabIndex        =   35
         Top             =   8910
         Width           =   1485
      End
      Begin FPSpread.vaSpread vasSch 
         Height          =   8595
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   14805
         _Version        =   393216
         _ExtentX        =   26114
         _ExtentY        =   15161
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         ColsFrozen      =   4
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   100
         OperationMode   =   2
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmInterface.frx":11A28
      End
   End
   Begin VB.CommandButton cmd_Trans 
      Caption         =   "전송"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10410
      TabIndex        =   43
      Top             =   9840
      Width           =   1485
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12000
      TabIndex        =   42
      Top             =   9840
      Width           =   1485
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   13590
      TabIndex        =   41
      Top             =   9840
      Width           =   1485
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   735
      Left            =   0
      TabIndex        =   37
      Top             =   3300
      Visible         =   0   'False
      Width           =   915
      _Version        =   393216
      _ExtentX        =   1614
      _ExtentY        =   1296
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
      SpreadDesigner  =   "frmInterface.frx":14779
   End
   Begin VB.CommandButton cmdSch 
      Caption         =   "검색"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10410
      TabIndex        =   33
      Top             =   180
      Width           =   1485
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   30
      Top             =   4650
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   30
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.TextBox txtTemp 
      Height          =   270
      Left            =   30
      TabIndex        =   7
      Top             =   870
      Visible         =   0   'False
      Width           =   705
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   8895
      Left            =   90
      TabIndex        =   0
      Top             =   870
      Width           =   8025
      _Version        =   393216
      _ExtentX        =   14155
      _ExtentY        =   15690
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      ColsFrozen      =   4
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      Protect         =   0   'False
      SpreadDesigner  =   "frmInterface.frx":149A0
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   585
      Left            =   90
      TabIndex        =   1
      Top             =   9780
      Width           =   15075
      _Version        =   65536
      _ExtentX        =   26591
      _ExtentY        =   1032
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelOuter      =   1
      Begin Threed.SSPanel sspPort 
         Height          =   465
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   5535
         _Version        =   65536
         _ExtentX        =   9763
         _ExtentY        =   820
         _StockProps     =   15
         Caption         =   "   CLINILOG 장비"
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Alignment       =   1
         Begin VB.Label lblCA 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "연결"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   2430
            TabIndex        =   9
            Top             =   120
            Width           =   360
         End
         Begin VB.Label lblCACom 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "[COM1]9600,n,8,1"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   2910
            TabIndex        =   8
            Top             =   120
            Width           =   1530
         End
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   735
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   15075
      _Version        =   65536
      _ExtentX        =   26591
      _ExtentY        =   1296
      _StockProps     =   15
      Caption         =   "     CLINILOG  Interface"
      ForeColor       =   8388736
      BackColor       =   16056319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   11940
         Picture         =   "frmInterface.frx":166AF
         ScaleHeight     =   255
         ScaleWidth      =   285
         TabIndex        =   67
         Top             =   60
         Width           =   315
      End
      Begin VB.TextBox txtUID 
         Appearance      =   0  '평면
         BackColor       =   &H00C0FFFF&
         Height          =   315
         Left            =   5520
         TabIndex        =   66
         Top             =   390
         Width           =   1785
      End
      Begin VB.TextBox txtBarcode 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8610
         TabIndex        =   39
         Top             =   180
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker dtpExamDate 
         Height          =   345
         Left            =   5520
         TabIndex        =   22
         Top             =   15
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   94830593
         CurrentDate     =   38584
      End
      Begin VB.CheckBox chkMode 
         Caption         =   "AUTO"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   585
         Left            =   9720
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   90
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtRemark 
         Height          =   435
         Left            =   2220
         TabIndex        =   4
         Top             =   1230
         Width           =   1545
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
         Caption         =   "사용자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12390
         TabIndex        =   68
         Top             =   90
         Width           =   1155
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검 사 자"
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
         Left            =   4530
         TabIndex        =   65
         Top             =   420
         Width           =   915
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "바코드번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7410
         TabIndex        =   40
         Top             =   255
         Width           =   1050
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
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
         Left            =   4530
         TabIndex        =   23
         Top             =   90
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "국립암센터 진단검사의학과"
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
         Left            =   11970
         TabIndex        =   21
         Top             =   480
         Width           =   2820
      End
   End
   Begin FPSpread.vaSpread vasRes 
      Height          =   7185
      Left            =   8160
      TabIndex        =   45
      Top             =   2580
      Width           =   7005
      _Version        =   393216
      _ExtentX        =   12356
      _ExtentY        =   12674
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      ColsFrozen      =   4
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":16C39
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1635
      Left            =   8160
      TabIndex        =   46
      Top             =   870
      Width           =   7005
      _Version        =   65536
      _ExtentX        =   12356
      _ExtentY        =   2884
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtEquip1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txtWkNo1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txtRack1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5790
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   180
         Width           =   1065
      End
      Begin VB.TextBox txtResDate1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtPName1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   615
         Width           =   1545
      End
      Begin VB.TextBox txtPID1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   615
         Width           =   1545
      End
      Begin VB.TextBox txtID1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   180
         Width           =   1545
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   405
         Left            =   5640
         TabIndex        =   62
         Top             =   570
         Width           =   1245
         _Version        =   65536
         _ExtentX        =   2196
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "출 력"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사장비"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   60
         Top             =   1140
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "W/L No"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3030
         TabIndex        =   58
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Rack-Pos"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4860
         TabIndex        =   57
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "결과시간"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3030
         TabIndex        =   56
         Top             =   1140
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "환자이름"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3030
         TabIndex        =   55
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "등록번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   54
         Top             =   675
         Width           =   840
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   53
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일"
      Begin VB.Menu subClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu subN1 
         Caption         =   "-"
      End
      Begin VB.Menu subClose 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnuPort 
      Caption         =   "연결"
      Begin VB.Menu subSendMode 
         Caption         =   "서버 결과 전송"
         Begin VB.Menu subSend1 
            Caption         =   "Auto"
         End
         Begin VB.Menu subSend2 
            Caption         =   "Manual"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "설정"
      Begin VB.Menu subCodeSet 
         Caption         =   "검사코드설정"
      End
      Begin VB.Menu subComSetup 
         Caption         =   "통신설정"
      End
      Begin VB.Menu subQCInfo 
         Caption         =   "QC 설정"
         Visible         =   0   'False
      End
      Begin VB.Menu subErrorCode 
         Caption         =   "Error Code"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "검색"
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gResCol As Long
Dim gCurRow As Long
Dim gMaxCol As Long

Dim iRow1 As Long
Dim iRow2 As Long
Dim iCol1 As Long
Dim iCol2 As Long

Dim SelVas As Integer

Private Type typeResult
    TestDate                As String
    TestTime                As String
    '
    FormatTypeCode          As String * 4
    TypeOfSample            As String * 2
    SampleID                As String * 20
    PatientID               As String * 10
    RackID                  As String * 10
    RackPosition            As String * 2
    NoOfAnalyzer            As String * 2
    '
    SampleInfCyle           As String * 4
    SampleInfHb             As String * 4
    SampleInfBil            As String * 4
    '
    NoOfItems               As String * 4
    ItemNo(1 To 100)        As String * 10
    Result(1 To 100)        As String * 10
    Comment(1 To 100)       As String * 4
    DilutionRatio(1 To 100) As String * 2
    ConfirmFlag(1 To 100)   As String * 2
    '
    LengthOfFreeComment     As String * 4
    FreeComment             As String * 32
End Type
Dim typResult As typeResult

Function GetSetup() As Boolean
'---------------------------------------------------------------------------------------------------------------------
'                       Setup  File을 읽어온다.
'---------------------------------------------------------------------------------------------------------------------
    Dim db_tmp As String * 100

    db_tmp = ""
    
    GetSetup = False
    
    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "driver", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Driver = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "uid", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.User = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "pwd", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Passwd = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "server", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.Server = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "database", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.DB = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("DATABASE", "hostname", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gDB_Parm.HostName = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("OPTION", "InsCode", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gInsCode = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerPath", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gServerPath = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "ServerID", "", db_tmp, 100, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gServerID = Trim(txtTemp)
    
    db_tmp = ""
    Call GetPrivateProfileString("Server", "사용자", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    gIFUser = Trim(txtTemp)
        
        
    GetSetup = True

End Function


Private Sub chkAll_Click()
    vasList.Row = -1
    vasList.Col = 1
    vasList.Value = chkAll.Value
End Sub

Private Sub chkMode_Click()
    If chkMode.Value = 1 Then
        chkMode.Caption = "Auto"
        SaveSetting "MEDIMATE", "CLINILOG", "SendMode", "1"
    Else
        chkMode.Caption = "Manual"
        SaveSetting "MEDIMATE", "CLINILOG", "SendMode", "0"
    End If
End Sub

Private Sub cmd_Trans_Click()
    Dim lRow, i, liEquipCode As Long
    Dim lsID As String
    Dim lsResult As String
    Dim lsWBC As String
    Dim lsNRBC As String
    Dim lsEOSIN As String
    Dim lsC_WBC As String
    Dim liRet As Integer
    Dim lsExamCode As String
    
    If MsgBox(" " & vbCrLf & "검사 결과를 전송하시겠습니까?" & vbCrLf & " ", vbInformation + vbYesNo + vbDefaultButton2, "결과 전송 알림") = vbNo Then
        Exit Sub
    End If
    
    For lRow = 1 To vasList.DataRowCnt
        vasList.Row = lRow
        vasList.Col = 1
        
        If vasList.Value = 1 Then
'            lsID = Trim(GetText(vasList, lRow, 2))
            
            
            liRet = 1
            
'            ClearSpread vasTemp
'
'            SQL = "Select barcode, examcode, examname, equipres, result, refflag "
'            SQL = SQL & " from pat_res " & vbCrLf & _
'                  "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
'                  "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'                  "  AND barcode = '" & Trim(lsID) & "' "
'            res = db_select_Vas(gLocal, SQL, vasTemp)
'
'            For liEquipCode = 1 To vasTemp.DataRowCnt
'                lsExamCode = Trim(GetText(vasTemp, liEquipCode, 2))
'                lsResult = Trim(GetText(vasTemp, liEquipCode, 5))
'
'                If lsExamCode <> "" And lsResult <> "" Then
'                    gInsCode = Left(SetEquip(Trim(GetText(vasList, lRow, 9))), 2)
'                    If Set_EqpResultsql(lsExamCode, lsResult, "", lsID, gInsCode) Then
'                    Else
'                        liRet = -1
'                    End If
'                End If
'            Next liEquipCode

            liRet = ToServer(lRow, vasList)
            If liRet = 1 Then
                SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                SetText vasList, "완료", lRow, gResCol
                
                vasList.Row = lRow
                vasList.Col = 1
                vasList.Value = 0
                
                Update_Sample Trim(GetText(vasList, lRow, 2))
                DeleteWorkList Trim(GetText(vasList, lRow, 2))
                
            Else
                SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                SetText vasList, "실패", lRow, gResCol
            End If
                        
        End If
    Next lRow
End Sub

Private Sub cmdClear_Click()
    subClear_Click
End Sub

Private Sub cmdCode_Click()
    frmCode.Show 1
End Sub

Private Sub cmdComSetup_Click()
    frmConfig.Show 1
End Sub
Private Sub cmdClose_Click()
    subClose_Click
End Sub

Private Sub cmdCloseDetail_Click()
    Frame1.Visible = False
End Sub

Private Sub cmdExcel_Click()
    Dim vasSelect As vaSpread
    Dim lRow, lCol, i As Long
    
    Dim xl As New Excel.Application
    Dim xlw As Excel.Workbook
    
    Dim IsInsName As String
    Dim lsFile As String
    
    Dim db_tmp As String * 20
    
    On Error GoTo ErrHandle
    
    db_tmp = ""
    Call GetPrivateProfileString("OPTION", "InsName", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    IsInsName = Trim(txtTemp)
    
    
    If frameSch.Visible = True Then
        Set vasSelect = vasSch
    Else
        Set vasSelect = vasList
    End If
    
    If iRow1 = iRow2 Then
        iRow1 = 1
        iRow2 = vasSelect.DataRowCnt
    End If
    
    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|All (*.*)|*.*"
    CommonDialog1.FileName = "C:\Documents and Settings\cdp\바탕 화면\" & Format(dtpExamDate.Value, "yyyymmdd") & "_" & IsInsName & ".xls"
    CommonDialog1.ShowOpen
    lsFile = CommonDialog1.FileName
    
    If Trim(lsFile) = "" Then
        MsgBox "파일이름이 없어 작업을 진행할 수 없습니다"
        Exit Sub
    End If
    
    Me.MousePointer = 11
    
    xl.Workbooks.Add
    Set xlw = xl.ActiveWorkbook
    
    xlw.Sheets("Sheet1").Select
    
    lRow = 0
    i = 5
    xlw.Application.Cells(lRow + 1, 2) = " " & Trim(GetText(vasSelect, lRow, 2)) & " "
    xlw.Application.Cells(lRow + 1, 3) = " " & Trim(GetText(vasSelect, lRow, 3)) & " "
    xlw.Application.Cells(lRow + 1, 4) = " " & Trim(GetText(vasSelect, lRow, 4)) & " "
    xlw.Application.Cells(lRow + 1, 5) = " " & Trim(GetText(vasSelect, lRow, 5)) & " "
    For lCol = gResCol + 1 To vasSelect.MaxCols - 1
        i = i + 1
        xlw.Application.Cells(1, i) = " " & Trim(GetText(vasSelect, lRow, lCol)) & " "
    Next lCol
    For lRow = iRow1 To iRow2
        xlw.Application.Cells(lRow + 1, 2) = " " & Trim(GetText(vasSelect, lRow, 2)) & " "
        xlw.Application.Cells(lRow + 1, 3) = " " & Trim(GetText(vasSelect, lRow, 3)) & " "
        xlw.Application.Cells(lRow + 1, 4) = " " & Trim(GetText(vasSelect, lRow, 4)) & " "
        xlw.Application.Cells(lRow + 1, 5) = " " & Trim(GetText(vasSelect, lRow, 5)) & " "
        i = 5
        For lCol = gResCol + 1 To vasSelect.MaxCols - 1
            i = i + 1
            xlw.Application.Cells(lRow + 1, i) = " " & Trim(GetText(vasSelect, lRow, lCol)) & " "
        Next lCol
    Next lRow
    If Dir(lsFile) <> "" Then
        Kill lsFile
    End If
    xlw.SaveAs lsFile
    xlw.Close
    
    Set xlw = Nothing
    Set xl = Nothing
    
    Me.MousePointer = 0
    
    Exit Sub
ErrHandle:
    Me.MousePointer = 0
    Exit Sub
End Sub

Private Sub cmdNext_Click()
    Dim argSpread As vaSpread
    Dim argRes As vaSpread
    
    Dim lRow1, lRow, lCol As Long
    
    If SelVas = 1 Then
        Set argSpread = vasList
    ElseIf SelVas = 2 Then
        Set argSpread = vasSch
    End If
    lRow1 = argSpread.ActiveRow
    lRow1 = lRow1 + 1
    
    vasActiveCell argSpread, lRow1, 2
    
    If lRow1 = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf lRow1 = argSpread.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If argSpread.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    txtID = Trim(GetText(argSpread, lRow1, 2))
    txtPID = Trim(GetText(argSpread, lRow1, 3))
    txtPName = Trim(GetText(argSpread, lRow1, 4))
    txtRack = Trim(GetText(argSpread, lRow1, 6))
    txtTube = Trim(GetText(argSpread, lRow1, 7))
    txtEquip = Trim(GetText(argSpread, lRow1, 5))
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    'lCol = gResCol
    lRow = 0
    For lCol = gResCol + 1 To gMaxCol - 1
        If Trim(GetText(argSpread, lRow1, lCol)) <> "" Then
            lRow = lRow + 1
            If lRow <= 20 Then
                Set argRes = vasRes1
            Else
                Set argRes = vasRes2
            End If
            If lRow = 21 Then lRow = 1
            
            SetText argRes, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argRes, Trim(GetText(argSpread, lRow1, lCol)), lRow, 3
            SetText argRes, Trim(GetText(argSpread, 0, lCol)), lRow, 2
            
            argSpread.Row = lRow1
            argSpread.Col = lCol
            Select Case argSpread.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argRes, lRow, lRow, 4, 4, 255, 127, 0
                SetText argRes, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argRes, lRow, lRow, 4, 4, 0, 127, 255
                SetText argRes, "▼", lRow, 4
            Case Else
                SetText argRes, "", lRow, 4
            End Select
        
        End If
    Next lCol

End Sub

Private Sub cmdPrev_Click()
    Dim argSpread As vaSpread
    Dim argRes As vaSpread
    
    Dim lRow1, lRow, lCol As Long
    
    If SelVas = 1 Then
        Set argSpread = vasList
    ElseIf SelVas = 2 Then
        Set argSpread = vasSch
    End If
    lRow1 = argSpread.ActiveRow
    lRow1 = lRow1 - 1
    
    vasActiveCell argSpread, lRow1, 2
    
    If lRow1 = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf lRow1 = argSpread.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If argSpread.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    txtID = Trim(GetText(argSpread, lRow1, 2))
    txtPID = Trim(GetText(argSpread, lRow1, 3))
    txtPName = Trim(GetText(argSpread, lRow1, 4))
    txtRack = Trim(GetText(argSpread, lRow1, 6))
    txtTube = Trim(GetText(argSpread, lRow1, 7))
    txtEquip = Trim(GetText(argSpread, lRow1, 5))
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    'lCol = gResCol
    lRow = 0
    For lCol = gResCol + 1 To gMaxCol - 1
        If Trim(GetText(argSpread, lRow1, lCol)) <> "" Then
            lRow = lRow + 1
            If lRow <= 20 Then
                Set argRes = vasRes1
            Else
                Set argRes = vasRes2
            End If
            If lRow = 21 Then lRow = 1
            
            SetText argRes, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argRes, Trim(GetText(argSpread, lRow1, lCol)), lRow, 3
            SetText argRes, Trim(GetText(argSpread, 0, lCol)), lRow, 2
            
            argSpread.Row = lRow1
            argSpread.Col = lCol
            Select Case argSpread.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argRes, lRow, lRow, 4, 4, 255, 127, 0
                SetText argRes, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argRes, lRow, lRow, 4, 4, 0, 127, 255
                SetText argRes, "▼", lRow, 4
            Case Else
                SetText argRes, "", lRow, 4
            End Select
        
        End If
    Next lCol

End Sub

Private Sub cmdPrint_Click()
    Dim sHead, sFoot As String
    Dim lRow, i As Long
    
    'Dim Width1, Width2
    
    On Error GoTo PrtErr
    
    'vasPrint.Visible = True
    
    vasPrint.SetText 1, 1, "  등록번호 : " & Trim(txtPID1)
    vasPrint.SetText 1, 2, "  성    명 : " & Trim(txtPName1)
    vasPrint.SetText 3, 2, "바코드 No : " & Trim(txtID1)
    vasPrint.SetText 4, 1, "    검사일자 : " & Format(CDate(txtResDate1), "yyyy.mm.dd")
        
    vasPrint.ClearRange 2, 5, 5, 44, True
    
    SQL = "Select '  ' & b.ExamName1, a.result, b.reflow & '-' & b.refhigh, b.unitcode " & vbCrLf
    SQL = SQL & "from pat_res a, equipexam b " & vbCrLf
    SQL = SQL & "where a.equipno = '" & gEquip & "' " & vbCrLf
    SQL = SQL & "  and a.examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf
    SQL = SQL & "  and a.barcode  = '" & Trim(txtID1) & "' " & vbCrLf
    SQL = SQL & "  and b.equip = a.equipno " & vbCrLf
    SQL = SQL & "  and b.EquipCode = a.EquipCode " & vbCrLf
    SQL = SQL & "  and b.examcode = a.examcode " & vbCrLf
    SQL = SQL & "order by b.seqno "
    'res = db_select_Vas(gLocal, SQL, vasPrint, 5, 2)
    
    Set cmdSQL.ActiveConnection = cn
    cmdSQL.CommandText = SQL
    Set rs = cmdSQL.Execute
  
    If rs.EOF = True Or rs.BOF = True Then
        Exit Sub
    End If
    
    lRow = 5
    While Not rs.EOF
        For i = 0 To rs.Fields.Count - 1
            vasPrint.Row = lRow
            vasPrint.Col = i + 2
            If IsNull(rs.Fields.Item(i).Value) Then
                vasPrint.Text = ""
            Else
                vasPrint.Text = CStr(rs.Fields.Item(i).Value)
            End If
        Next i
        rs.MoveNext
        lRow = lRow + 1
    Wend
    i = lRow - 1
    'vasPrint.Visible = True
    
'    i = 4
'    For lRow = 1 To vasRes1.DataRowCnt
'        If Trim(GetText(vasRes1, lRow, 1)) <> "" Then
'            SQL = "select equipcode, reflow, refhigh, unitcode, ExamName1 from equipexam " & vbCrLf & _
'                  "Where Equip = '" & gEquip & "' " & CR & _
'                  "  and EquipCode = '" & Trim(GetText(vasRes1, lRow, 1)) & "' "
'            res = db_select_Col(gLocal, SQL)
'
'            i = i + 1
'            If Trim(gReadBuf(4)) = "" Then
'                vasPrint.SetText 2, i, "  " & Trim(GetText(vasRes1, lRow, 2))
'            Else
'                vasPrint.SetText 2, i, "  " & Trim(gReadBuf(4))
'            End If
'            vasPrint.SetText 3, i, Trim(GetText(vasRes1, lRow, 3))
'            If Trim(gReadBuf(0)) = Trim(GetText(vasRes1, lRow, 1)) Then
'                If Trim(gReadBuf(1)) <> "" And Trim(gReadBuf(2)) <> "" Then
'                    vasPrint.SetText 4, i, Trim(gReadBuf(1)) & " - " & Trim(gReadBuf(2))
'                End If
'                vasPrint.SetText 5, i, Trim(gReadBuf(3))
'            Else
'                vasPrint.SetText 3, i, ""
'                vasPrint.SetText 4, i, ""
'            End If
'        End If
'    Next lRow
'    For lRow = 1 To vasRes2.DataRowCnt
'        If Trim(GetText(vasRes2, lRow, 1)) <> "" Then
'            SQL = "select equipcode, reflow, refhigh, unitcode, ExamName1 from equipexam " & vbCrLf & _
'                  "Where Equip = '" & gEquip & "' " & CR & _
'                  "  and EquipCode = '" & Trim(GetText(vasRes2, lRow, 1)) & "' "
'            res = db_select_Col(gLocal, SQL)
'
'            i = i + 1
'            If Trim(gReadBuf(4)) = "" Then
'                vasPrint.SetText 2, i, "  " & Trim(GetText(vasRes2, lRow, 2))
'            Else
'                vasPrint.SetText 2, i, "  " & Trim(gReadBuf(4))
'            End If
'            vasPrint.SetText 3, i, Trim(GetText(vasRes2, lRow, 3))
'            If Trim(gReadBuf(0)) = Trim(GetText(vasRes2, lRow, 1)) Then
'                If Trim(gReadBuf(1)) <> "" And Trim(gReadBuf(2)) <> "" Then
'                    vasPrint.SetText 4, i, Trim(gReadBuf(1)) & " - " & Trim(gReadBuf(2))
'                End If
'                vasPrint.SetText 5, i, Trim(gReadBuf(3))
'            Else
'                vasPrint.SetText 3, i, ""
'                vasPrint.SetText 4, i, ""
'            End If
'        End If
'    Next lRow
    
    sHead = "/fn""굴림체"" /fz""18"" /fb1 /fi0 /fu0 " & "/c" & "생 화 학 결 과 지" & "/n/n "
    sFoot = "/fn""굴림체"" /fz""9"" /fb0 /fi0 /fu0 " & "/c" & "진단검사의학과 생화학실" & "/n/n"
    
    vasPrint.PrintOrientation = 1
    vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasPrint.PrintJobName = "생화학결과지-TBA(" & gInsCode & ")"
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot
    vasPrint.PrintMarginTop = 1200
    vasPrint.PrintMarginBottom = 600
    vasPrint.PrintMarginLeft = 720
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = False
    
    'Set printing range
    vasPrint.Row = 1
    vasPrint.Row2 = i
    vasPrint.Col = 1
    vasPrint.Col2 = 5
    vasPrint.PrintType = PrintTypeCellRange
    
    vasPrint.PrintType = PrintTypeCellRange 'SS_PRINT_ALL(default)
    
    
    'vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT
    
    Exit Sub
PrtErr:
    MsgBox "프린터 오류!"
    Exit Sub
End Sub

Private Sub cmdSch_Click()
    Dim lRow, lCol As Long
    Dim lsID As String
    Dim liEquipCode As Integer
    
    Dim rs_Res As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    ClearSpread vasList
    
    SQL = "Select distinct a.barcode, a.pid, a.pname, a.examtype, a.diskno, " & _
            "a.posno, a.sendflag, a.examuid, a.barcode  " & vbCrLf & _
          "from pat_res a" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " '& vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.receno = 'Q' " & vbCrLf & _
          "Order by a.barcode, a.equipcode "
    res = db_select_Vas(gLocal, SQL, vasList, 1, 2)
    vasList_Click 2, 1
    
    Exit Sub

    ClearSpread vasSch, 0, 2
    
    'frameSch.Visible = True
    
    Me.MousePointer = 11
    
    vasSch.MaxCols = vasList.MaxCols
    
    For lCol = 1 To vasList.MaxCols
        SetText vasSch, Trim(GetText(vasList, 0, lCol)), 0, lCol
        vasSch.ColWidth(lCol) = vasList.ColWidth(lCol)
    Next lCol
    
    SQL = "Select a.barcode, a.pid, a.pname, a.examtype, a.diskno, " & _
            "a.posno, a.sendflag, a.equipcode, a.examcode, a.result, " & _
            "a.refflag, a.resdate, a.seqno  " & vbCrLf & _
          "from pat_res a" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.receno = 'Q' " & vbCrLf & _
          "Order by a.barcode, a.equipcode "
    
    Set rs_Res = db_select_rs(gLocal, SQL)
       
    If rs_Res Is Nothing Then
    Else
        lsID = "검체번호"
        lRow = 0
        Do While Not rs_Res.EOF
            If Trim(CStr(rs_Res.Fields.Item(0).Value)) <> lsID Then
                lRow = lRow + 1
                
                If lRow > vasSch.MaxRows Then
                    vasSch.MaxRows = lRow
                End If
                
                For lCol = 2 To 8
                    SetText vasSch, Trim(CStr(rs_Res.Fields.Item(lCol - 2).Value)), lRow, lCol
                Next lCol
                
                vasSch.SetText gMaxCol, lRow, Trim(CStr(rs_Res.Fields.Item(0).Value))
            End If
            
            If Trim(rs_Res.Fields.Item(7).Value) <> "" Then
                For liEquipCode = 1 To UBound(gArrExam)
                    If CInt(gArrExam(liEquipCode, 1)) = Trim(rs_Res.Fields.Item(7).Value) Then
                        lCol = gResCol + liEquipCode
                        'lCol = liEquipCode - 1
        '                If liEquipCode = 1 Then
        '                    MsgBox ""
        '                End If
                        SetText vasSch, Trim(CStr(rs_Res.Fields.Item(9).Value)), lRow, lCol
                        Select Case Trim(CStr(rs_Res.Fields.Item(10).Value))
                        Case "H"
                            SetForeColor vasSch, lRow, lRow, lCol, lCol, 255, 127, 0
                        Case "L"
                            SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 127, 255
                        Case Else
                            SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 0, 0
                        End Select
                        
                        Exit For
                    End If
                Next liEquipCode
            End If
            
            lsID = Trim(CStr(rs_Res.Fields.Item(0).Value))
            
            rs_Res.MoveNext
        Loop
        
        rs_Res.Close
    End If
    
    SQL = "Select a.barcode, a.pid, a.pname, a.examtype, a.diskno, " & _
            "a.posno, a.sendflag, a.equipcode, a.examcode, a.result, " & _
            "a.refflag, a.resdate, a.seqno  " & vbCrLf & _
          "from pat_res a" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.receno <> 'Q' " & vbCrLf & _
          "Order by a.barcode, a.equipcode "
    
    Set rs_Res = db_select_rs(gLocal, SQL)
       
    If rs_Res Is Nothing Then
    Else
    
        lsID = ""
        lRow = vasSch.DataRowCnt + 1
        
        Do While Not rs_Res.EOF
            If Trim(CStr(rs_Res.Fields.Item(0).Value)) <> lsID Then
                lRow = lRow + 1
                
                If lRow > vasSch.MaxRows Then
                    vasSch.MaxRows = lRow
                End If
                
                For lCol = 2 To 8
                    SetText vasSch, Trim(CStr(rs_Res.Fields.Item(lCol - 2).Value)), lRow, lCol
                Next lCol
                
                vasSch.SetText gMaxCol, lRow, Trim(CStr(rs_Res.Fields.Item(0).Value))
            End If
            
            If Trim(rs_Res.Fields.Item(7).Value) <> "" Then
                For liEquipCode = 1 To UBound(gArrExam)
                    If CInt(gArrExam(liEquipCode, 1)) = Trim(rs_Res.Fields.Item(7).Value) Then
                        lCol = gResCol + liEquipCode
                        'lCol = liEquipCode - 1
        '                If liEquipCode = 1 Then
        '                    MsgBox ""
        '                End If
                        SetText vasSch, Trim(CStr(rs_Res.Fields.Item(9).Value)), lRow, lCol
                        Select Case Trim(CStr(rs_Res.Fields.Item(10).Value))
                        Case "H"
                            SetForeColor vasSch, lRow, lRow, lCol, lCol, 255, 127, 0
                        Case "L"
                            SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 127, 255
                        Case Else
                            SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 0, 0
                        End Select
                        
                        Exit For
                    End If
                Next liEquipCode
            End If
            
            lsID = Trim(CStr(rs_Res.Fields.Item(0).Value))
            
            rs_Res.MoveNext
        Loop
        
        rs_Res.Close
    End If
    
    Me.MousePointer = 0
    
    vasSch.RowHeight(-1) = 14.7
    
    vasList.Row = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col = gResCol + 1
    vasList.Col2 = vasList.MaxCols
    vasList.BlockMode = True
    vasList.CellType = CellTypeEdit
    vasList.TypeHAlign = TypeHAlignCenter
    vasList.TypeVAlign = TypeVAlignCenter
    vasList.BlockMode = False
    
    frameSch.Visible = True
    'EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh
    
    Exit Sub
ErrHandle:
    Me.MousePointer = 0
    Exit Sub

End Sub

Sub SearchSample(ByVal asBarcode As String)
    Dim lRow, lCol As Long
    Dim lsID As String
    Dim liEquipCode As Integer
    
    Dim rs_Res As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    vasSch.Row = -1
    vasSch.Col = 1
    vasSch.Value = 0
    
    ClearSpread vasSch, 0, 2
    
    'frameSch.Visible = True
    
    Me.MousePointer = 11
    
    vasSch.MaxCols = vasList.MaxCols
    
    For lCol = 1 To vasList.MaxCols
        SetText vasSch, Trim(GetText(vasList, 0, lCol)), 0, lCol
        vasSch.ColWidth(lCol) = vasList.ColWidth(lCol)
    Next lCol
    
    SQL = "Select a.barcode, a.pid, a.pname, a.examtype, a.diskno, " & _
            "a.posno, a.sendflag, a.equipcode, a.examcode, a.result, " & _
            "a.refflag, a.resdate, a.seqno  " & vbCrLf & _
          "from pat_res a" & vbCrLf & _
          "where a.examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & asBarcode & "' " & vbCrLf & _
          "Order by a.barcode, a.equipcode "
    
    Set rs_Res = db_select_rs(gLocal, SQL)
       
    If rs_Res Is Nothing Then GoTo ErrHandle
    
    lsID = ""
    lRow = 0
    
    Do While Not rs_Res.EOF
        If Trim(CStr(rs_Res.Fields.Item(0).Value)) <> lsID Then
            lRow = lRow + 1
            
            If lRow > vasSch.MaxRows Then
                vasSch.MaxRows = lRow
            End If
            
            For lCol = 2 To 8
                SetText vasSch, Trim(CStr(rs_Res.Fields.Item(lCol - 2).Value)), lRow, lCol
            Next lCol
            
            vasSch.SetText gMaxCol, lRow, Trim(CStr(rs_Res.Fields.Item(0).Value))
        End If
        
        If Trim(rs_Res.Fields.Item(7).Value) <> "" Then
            For liEquipCode = 1 To UBound(gArrExam)
                If CInt(gArrExam(liEquipCode, 1)) = Trim(rs_Res.Fields.Item(7).Value) Then
                    lCol = gResCol + liEquipCode
                    'lCol = liEquipCode - 1
    '                If liEquipCode = 1 Then
    '                    MsgBox ""
    '                End If
                    SetText vasSch, Trim(CStr(rs_Res.Fields.Item(9).Value)), lRow, lCol
                    Select Case Trim(CStr(rs_Res.Fields.Item(10).Value))
                    Case "H"
                        SetForeColor vasSch, lRow, lRow, lCol, lCol, 255, 127, 0
                    Case "L"
                        SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 127, 255
                    Case Else
                        SetForeColor vasSch, lRow, lRow, lCol, lCol, 0, 0, 0
                    End Select
                    
                    Exit For
                End If
            Next liEquipCode
        End If
        
        lsID = Trim(CStr(rs_Res.Fields.Item(0).Value))
        
        rs_Res.MoveNext
    Loop
    
    rs_Res.Close
    
    Me.MousePointer = 0
    
    vasSch.RowHeight(-1) = 14.7
    
    frameSch.Visible = True
    'EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh
    
    Exit Sub
ErrHandle:
    Me.MousePointer = 0
    Exit Sub

End Sub

Private Sub cmdSchClose_Click()
    frameSch.Visible = False
End Sub

Private Sub Command1_Click()
    CLINILOG
    txtBuff = ""
End Sub

Private Sub Form_Load()
    Dim lsTmp As String * 20
    Dim lInt As Long
    
    Me.Left = 0
    Me.Top = 0
    
    dtpExamDate.Value = Format(Date, "yyyy-mm-dd")
'
    gResCol = 8
    
    'frmConnect.Show 1
    GetComSetup
    
    MSComm1.CommPort = CA_COM.ComPort
    MSComm1.Settings = CA_COM.Speed & "," & CA_COM.Parity & "," & CA_COM.DataBit & "," & CA_COM.StartBit
    If CA_COM.RTSEnable = "1" Then
        MSComm1.RTSEnable = True
    Else
        MSComm1.RTSEnable = False
    End If
    If CA_COM.DTREnable = "1" Then
        MSComm1.DTREnable = True
    Else
        MSComm1.DTREnable = False
    End If
    MSComm1.PortOpen = True
    
    lblCACom.Caption = "[COM" & CA_COM.ComPort & "]" & MSComm1.Settings
    
    cn_Local_Flag = False
    cn_Server_Flag = False
    
    GetSetup
    
    lblUser.Caption = Trim(gIFUser)
    
    If Connect_Local Then
        cn_Local_Flag = True
    End If
    
'    If Connect_Server Then
'        cn_Server_Flag = True
'    End If
    
    '2009.06.25 윤영기=======================================
    SQL = "select * from errorcode "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "create table errorcode ( " & vbCrLf
        SQL = SQL & " tlaerror varchar(20), " & vbCrLf
        SQL = SQL & " chgerror varchar(20), " & vbCrLf
        SQL = SQL & " errordesc varchar(50) ) "
        res = SendQuery(gLocal, SQL)
    End If
    '========================================================
    
    SQL = "select * from qcinfo "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "create table qcinfo ( " & vbCrLf
        SQL = SQL & " inscode varchar(10), " & vbCrLf
        SQL = SQL & " insname varchar(50), " & vbCrLf
        SQL = SQL & " equipqc varchar(50), " & vbCrLf
        SQL = SQL & " qcid varchar(50) ) " & vbCrLf
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "select lowlimit from equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "alter table equipexam " & vbCrLf
        SQL = SQL & " add column lowlimit varchar(10)  " & vbCrLf
        res = SendQuery(gLocal, SQL)
    
        SQL = "alter table equipexam " & vbCrLf
        SQL = SQL & " add column highlimit varchar(10)  " & vbCrLf
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "select TiterValue from equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "alter table equipexam " & vbCrLf
        SQL = SQL & " add column TiterValue varchar(10)  " & vbCrLf
        res = SendQuery(gLocal, SQL)
    
        SQL = "alter table equipexam " & vbCrLf
        SQL = SQL & " add column TiterEqual  Integer  " & vbCrLf
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "select equipres from pat_res "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = "alter table pat_res " & vbCrLf
        SQL = SQL & " add column equipres varchar(50)  " & vbCrLf
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = "Select OrdCode from EquipExam "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SQL = "Alter table EquipExam add column OrdCode varchar(20) "
        res = SendQuery(gLocal, SQL)
        
        SQL = "Update EquipExam set OrdCode = ExamCode "
        res = SendQuery(gLocal, SQL)
    End If
    
    '2010.03.19 이상은
    SQL = " Alter Table pat_res Alter Column result text(150) "
    res = SendQuery(gLocal, SQL)
    
    If Trim(GetSetting("MEDIMATE", "CLINILOG", "SendMode", "0")) = "1" Then
        chkMode.Value = 1
        subSend1.Checked = True
        subSend2.Checked = False
    Else
        chkMode.Value = 0
        subSend1.Checked = False
        subSend2.Checked = True
    End If
    
    'GetExamCode
    
    txtBuff = ""
    
    ClearSpread vasList
    ClearSpread vasRes
    vasList.MaxRows = 1
    vasRes.MaxRows = 1
    
    'subClear_Click
    lsTmp = ""
    Call GetPrivateProfileString("DATA", "Days", "", lsTmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(lsTmp)
    If IsNumeric(Trim(txtTemp)) Then
        lInt = CLng(Trim(txtTemp))
    Else
        lInt = 30
    End If
    
    SQL = "Delete from pat_res " & vbCrLf & _
          "WHERE examdate < '" & Format(DateAdd("d", 0 - lInt, CDate(dtpExamDate.Value)), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' "
    'res = SendQuery(gLocal, SQL)
    
End Sub

Sub GetExamCode()
    Dim AdoRs_Exam As ADODB.Recordset
    Dim lCol As Long
    Dim i As Integer
    
'    ReDim gArrExam(0)
'    gArrExam(0) = ""
    
    'SQL = "SELECT EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh, RSGubun " & CR & _
          "  From EquipExam " & CR & _
          " WHERE Equip = '" & gEquip & "' " & CR & _
          "   and UseFlag = 1 " & vbCrLf & _
          " Order by seqno "
    SQL = "SELECT EquipCode, ExamName, SeqNo, count(ExamName) " & vbCrLf & _
          "  From EquipExam " & vbCrLf & _
          " WHERE Equip = '" & gEquip & "' " & vbCrLf & _
          "   and UseFlag = 1 " & vbCrLf & _
          " Group by EquipCode, ExamName, SeqNo " & vbCrLf & _
          " Order by SeqNo"
    
    Set AdoRs_Exam = db_select_rs(gLocal, SQL)
    If AdoRs_Exam Is Nothing Then
        ClearSpread vasList, 1, 1
    Else
        ClearSpread vasList, 1, 1
        
        AdoRs_Exam.MoveFirst
        lCol = gResCol
        
        
        Do Until AdoRs_Exam.EOF
            lCol = lCol + 1
             
'            ReDim Preserve gArrExam(lCol - gResCol, 5)
'
'            gArrExam(lCol - gResCol, 1) = AdoRs_Exam.Fields(0).Value
'            gArrExam(lCol - gResCol, 2) = AdoRs_Exam.Fields(1).Value
'            gArrExam(lCol - gResCol, 3) = AdoRs_Exam.Fields(2).Value
'            gArrExam(lCol - gResCol, 4) = AdoRs_Exam.Fields(3).Value
'            gArrExam(lCol - gResCol, 5) = AdoRs_Exam.Fields(4).Value
'
'            SetText vasList, AdoRs_Exam.Fields(2).Value, 0, lCol
'
            AdoRs_Exam.MoveNext
        Loop
        
        ReDim gArrExam(lCol - gResCol, 2)
        
        vasList.MaxCols = lCol + 1
        
        AdoRs_Exam.MoveFirst
        lCol = gResCol
        
        
        
        Do Until AdoRs_Exam.EOF
            lCol = lCol + 1
             
            'ReDim Preserve gArrExam(lCol - gResCol, 5)
            For i = 0 To 1
                If IsNull(AdoRs_Exam.Fields(i).Value) Then
                    gArrExam(lCol - gResCol, i + 1) = ""
                Else
                    gArrExam(lCol - gResCol, i + 1) = AdoRs_Exam.Fields(i).Value
                End If
            Next i
            
            SetText vasList, AdoRs_Exam.Fields(1).Value, 0, lCol
            vasList.Row = -1
            vasList.Col = lCol
            vasList.TypeHAlign = TypeHAlignCenter
            vasList.TypeVAlign = TypeVAlignCenter
            
            'vasActiveCell vasList, 0, lCol
            'Debug.Print Format(lCol, "00") & " : " & AdoRs_Exam.Fields(1).Value
            AdoRs_Exam.MoveNext
        Loop
    End If
    
    gMaxCol = lCol + 1
    
    vasList.MaxCols = gMaxCol
    
    vasList.ColWidth(gMaxCol) = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox(" " & vbCrLf & "   장비와 통신이 끊어져 데이타 전송이 중단됩니다    " & vbCrLf & _
              " " & vbCrLf & "   종료하시겠습니까?   " & vbCrLf & _
              " ", vbInformation + vbYesNo + vbDefaultButton2, "알림 : 종료 ") = vbNo Then
            
        Cancel = 1
    Else
        Cancel = 0
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisConnect_Local
    DisConnect_Server
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
End Sub

Private Sub MSComm1_OnComm()
    Dim lsChar As String
    
    lsChar = MSComm1.Input
    
    Select Case lsChar
    Case chrENQ
        If dtpExamDate.Value <> Format(Date, "yyyy-mm-dd") Then
            dtpExamDate.Value = Format(Date, "yyyy-mm-dd")
            
            
        End If
        
        SaveRes "[RX]<ENQ>"
        
        MSComm1.Output = chrACK
        SaveRes "[TX]<ACK>"
    Case chrEOT
        SaveRes "[RX]<EOT>"
        
    Case chrSTX
        
        txtBuff.Text = ""
        
        txtBuff.Text = txtBuff.Text & lsChar
    Case chrETX, chrETB
        txtBuff.Text = txtBuff.Text & lsChar
        
        SaveRes "[RX]" & txtBuff.Text
        
        CLINILOG
        
        MSComm1.Output = chrACK
        SaveRes "[TX]<ACK>"
        
    Case Else
        txtBuff.Text = txtBuff.Text & lsChar
    End Select
End Sub

Sub CLINILOG()
    Dim myVar As String
    Dim lsTmp As String
    Dim lsSampleInfo As String
    
    Dim iPoint As Integer
    
    Dim lsID As String
    Dim lsDate As String
    Dim lsRack As String
    Dim lsTube As String
    Dim lsSeqNo As String
    Dim lsGubun As String
    
    Dim SampleJudg, PosDif, PosMor, PosCnt, ErrFun, ErrRes As String
    Dim InfoOrd, InfoSample, InfoUnit, InfoWBC, InfoPLT As String
    
    Dim liEquipCode As Integer
    Dim lsEquipCode As String
    Dim lsResult As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsErr As String
    
    Dim lRow As Long
    Dim lResRow As Long
    Dim lCol As Long
    Dim i, j, k, z As Long
    
    Dim mExam As Variant
    
    Dim liRet As Integer
    
    Dim lsHead As String
    Dim lsOrd As String
    Dim lsOrder() As String
    
    Dim ReceiveData As String
    Dim Sum, ii
    
    Dim sIndexFlag As String
    
    Dim lsErrDesc As String
    
    ClearSpread vasRes
    
    sIndexFlag = ""
    
    If Trim(txtBuff) = "" Then Exit Sub
    
    ReceiveData = Mid(txtBuff, 2, Len(txtBuff) - 2)
    
    Sum = 1
    typResult.FormatTypeCode = Mid(ReceiveData, Sum, Len(typResult.FormatTypeCode))
    Sum = Sum + Len(typResult.FormatTypeCode)
    typResult.TypeOfSample = Mid(ReceiveData, Sum, Len(typResult.TypeOfSample))
    Sum = Sum + Len(typResult.TypeOfSample)
    typResult.SampleID = Mid(ReceiveData, Sum, Len(typResult.SampleID))
    Sum = Sum + Len(typResult.SampleID)
    'typResult.PatientID = Mid(ReceiveData, Sum, Len(typResult.PatientID))
    'Sum = Sum + Len(typResult.PatientID)
    typResult.RackID = Mid(ReceiveData, Sum, Len(typResult.RackID))
    Sum = Sum + Len(typResult.RackID)
    typResult.RackPosition = Mid(ReceiveData, Sum, Len(typResult.RackPosition))
    Sum = Sum + Len(typResult.RackPosition)
    typResult.NoOfAnalyzer = Mid(ReceiveData, Sum, Len(typResult.NoOfAnalyzer))
    Sum = Sum + Len(typResult.NoOfAnalyzer)
    '
    typResult.SampleInfCyle = Mid(ReceiveData, Sum, Len(typResult.SampleInfCyle))
    Sum = Sum + Len(typResult.SampleInfCyle)
    typResult.SampleInfHb = Mid(ReceiveData, Sum, Len(typResult.SampleInfHb))
    Sum = Sum + Len(typResult.SampleInfHb)
    typResult.SampleInfBil = Mid(ReceiveData, Sum, Len(typResult.SampleInfBil))
    Sum = Sum + Len(typResult.SampleInfBil)
    '
    typResult.NoOfItems = Mid(ReceiveData, Sum, Len(typResult.NoOfItems))
    Sum = Sum + Len(typResult.NoOfItems)
        For ii = 1 To Val(typResult.NoOfItems)
            typResult.ItemNo(ii) = Mid(ReceiveData, Sum, Len(typResult.ItemNo(ii)))
            Sum = Sum + Len(typResult.ItemNo(ii))
            typResult.Result(ii) = Mid(ReceiveData, Sum, Len(typResult.Result(ii)))
            Sum = Sum + Len(typResult.Result(ii))
            typResult.Comment(ii) = Mid(ReceiveData, Sum, Len(typResult.Comment(ii)))
            Sum = Sum + Len(typResult.Comment(ii))
            typResult.DilutionRatio(ii) = Mid(ReceiveData, Sum, Len(typResult.DilutionRatio(ii)))
            Sum = Sum + Len(typResult.DilutionRatio(ii))
            typResult.ConfirmFlag(ii) = Mid(ReceiveData, Sum, Len(typResult.ConfirmFlag(ii)))
            Sum = Sum + Len(typResult.ConfirmFlag(ii))
            
            If Trim(typResult.NoOfAnalyzer) = "61" Then 'Modular
                Select Case Trim(typResult.Comment(ii))
                Case "26", "44"
                    typResult.Result(ii) = ">" & Trim(typResult.Result(ii))
                Case "27", "45"
                    typResult.Result(ii) = "<" & Trim(typResult.Result(ii))
                End Select
            End If
            
            If Trim(typResult.ItemNo(ii)) = "LHI01" Or Trim(typResult.ItemNo(ii)) = "BLHI01" Then
                If IsNumeric(typResult.Result(ii)) Then
                    If CInt(typResult.Result(ii)) > 0 Then
                        sIndexFlag = sIndexFlag & " L" & Trim(typResult.Result(ii))
                    End If
                End If
            End If
            If Trim(typResult.ItemNo(ii)) = "LHI02" Or Trim(typResult.ItemNo(ii)) = "BLHI02" Then
                If IsNumeric(typResult.Result(ii)) Then
                    If CInt(typResult.Result(ii)) > 0 Then
                        sIndexFlag = sIndexFlag & " H" & Trim(typResult.Result(ii))
                    End If
                End If
            End If
            If Trim(typResult.ItemNo(ii)) = "LHI03" Or Trim(typResult.ItemNo(ii)) = "BLHI03" Then
                If IsNumeric(typResult.Result(ii)) Then
                    If CInt(typResult.Result(ii)) > 0 Then
                        sIndexFlag = sIndexFlag & " I" & Trim(typResult.Result(ii))
                    End If
                End If
            End If
        Next ii
    '
    typResult.LengthOfFreeComment = Mid(ReceiveData, Sum, Len(typResult.LengthOfFreeComment))
    Sum = Sum + Len(typResult.LengthOfFreeComment)
    typResult.FreeComment = Mid(ReceiveData, Sum, Len(typResult.FreeComment))
    Sum = Sum + Len(typResult.FreeComment)

    ReDim lsOrder(0)
    
    If IsNumeric(typResult.RackID) Then
        typResult.RackID = CStr(CCur(typResult.RackID))
    End If
    
    lRow = -1
    For i = 1 To vasList.DataRowCnt
        If Trim(typResult.SampleID) = Trim(GetText(vasList, i, 2)) Then
            lRow = i
            Exit For
        End If
    Next i
    
    If lRow = -1 Then
        lRow = vasList.DataRowCnt + 1
        If lRow > vasList.MaxRows Then vasList.MaxRows = lRow
    End If
    
    SetText vasList, Trim(typResult.SampleID), lRow, 2
    SetText vasList, Trim(typResult.RackID), lRow, 6
    SetText vasList, Trim(typResult.RackPosition), lRow, 7
    SetText vasList, Trim(typResult.SampleID), lRow, gResCol + 2
'    mExam = Get_OrderBody(Trim(typResult.SampleID))
'    If Not IsNull(mExam) Then
'        SetText vasList, Trim(mExam(1, LBound(mExam, 2))), lRow, 3
'        SetText vasList, Trim(mExam(2, LBound(mExam, 2))), lRow, 4
'
'        SetText vasList, Trim(mExam(5, LBound(mExam, 2))) & "-" & Trim(mExam(6, LBound(mExam, 2))), lRow, 5
'
'        SetText vasList, "결과", lRow, gResCol
'        SetBackColor vasList, lRow, lRow, 1, 1, 255, 250, 205
'    Else
'        vasList.Row = lRow
'        vasList.Col = 1
'        vasList.Value = 1
'        SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
'        SetText vasList, "오류", lRow, gResCol
'    End If

    res = Online_XML(gXml_S03, Trim(typResult.SampleID))
    If res = 1 Then
        SetText vasList, Trim(gPat_Info_Select.PT_NO), lRow, 3
        SetText vasList, Trim(gPat_Info_Select.PT_NM), lRow, 4
        
        SetText vasList, Trim(gPat_Info_Select.ACPTNO_1), lRow, 5
    
        SetText vasList, "결과", lRow, gResCol
        SetBackColor vasList, lRow, lRow, 1, 1, 255, 250, 205
    Else
        vasList.Row = lRow
        vasList.Col = 1
        vasList.Value = 1
        SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
        SetText vasList, "오류", lRow, gResCol
    End If
    
    SetText vasList, Trim(typResult.NoOfAnalyzer), lRow, gResCol + 1
    
        'ReDim gArrExamRes(1 To gMaxCol - 1 - gResCol)
    ReDim gArrExamRes(1 To 1)
        
    For ii = 1 To Val(typResult.NoOfItems)
        If Trim(typResult.ItemNo(ii)) <> "" And Trim(typResult.Result(ii)) <> "" Then
            If typResult.TypeOfSample = "00" Or typResult.TypeOfSample = "01" Then '00:normal, 01:control
                
                If typResult.TypeOfSample = "01" Then
                    lsGubun = "Q"
                Else
                    lsGubun = "P"
                End If
                
                lsEquipCode = Trim(typResult.ItemNo(ii))
                lsExamCode = Trim(typResult.ItemNo(ii))
                lsResult = Trim(typResult.Result(ii))
                
                
                ClearSpread vasTemp
                SQL = "SELECT EquipCode,ExamCode, ExamName, Seqno, PointSize, " & _
                      " RefLow, RefHigh, RSGubun, CutOffFlag, NegValue, " & _
                      " NegEqual, PosValue, PosEqual, TiterValue, TiterEqual,  " & vbCrLf & _
                      " lowlimit, highlimit " & vbCrLf & _
                      "  From EquipExam " & vbCrLf & _
                      " WHERE Equip = '" & gEquip & "' "

                SQL = SQL & vbCrLf & _
                      "   and ExamCode = '" & lsEquipCode & "' "

                SQL = SQL & vbCrLf & _
                      " Order by seqno, equipcode, examcode "
                res = db_select_Vas(gLocal, SQL, vasTemp)
                If vasTemp.DataRowCnt > 0 Then
                    
                    liEquipCode = liEquipCode + 1
                    
                    ReDim Preserve gArrExamRes(1 To liEquipCode)
                    
                    i = 1
                    
                    gArrExamRes(liEquipCode).EquipCode = Trim(GetText(vasTemp, i, 1))
                    gArrExamRes(liEquipCode).ExamCode = Trim(GetText(vasTemp, i, 2))
                    gArrExamRes(liEquipCode).ExamName = Trim(GetText(vasTemp, i, 3))
                    gArrExamRes(liEquipCode).SeqNo = Trim(GetText(vasTemp, i, 4))
                    gArrExamRes(liEquipCode).RefLow = Trim(GetText(vasTemp, i, 6))
                    gArrExamRes(liEquipCode).RefHigh = Trim(GetText(vasTemp, i, 7))
                    gArrExamRes(liEquipCode).RefFlag = ""
                    gArrExamRes(liEquipCode).res = lsResult
                    gArrExamRes(liEquipCode).EquipRes = lsResult
                    gArrExamRes(liEquipCode).EquipGubun = Trim(typResult.NoOfAnalyzer)
                    If Trim(typResult.Comment(ii)) = "00" Then
                        gArrExamRes(liEquipCode).Equipcomment = ""
                    Else
                        gArrExamRes(liEquipCode).Equipcomment = typResult.Comment(ii)
                    End If
                    
                    If liEquipCode = 1 And Trim(sIndexFlag) <> "" Then
                        gArrExamRes(liEquipCode).Equipcomment = Trim(gArrExamRes(liEquipCode).Equipcomment & " " & sIndexFlag)
                    End If
                    
                    SetResult1 liEquipCode, i
                    SetLimit liEquipCode
                    
'2010.04.02 이상은 - EMR에서 처리하기로 함
'                    Select Case gArrExamRes(liEquipCode).EquipCode
'                    Case "L5111"        'HBsAg
'                        If gArrExamRes(liEquipCode).res = "Positive" Then
'                            gArrExamRes(liEquipCode).res = "Positive(Pt''s Result : > 250.00, Cut-off:≥0.05)"
'                        End If
'                    Case "L5118"        'Anti-HCV
'                        If gArrExamRes(liEquipCode).res = "Positive" Then
'                            gArrExamRes(liEquipCode).res = "Positive(Pt''s Index : > 11.00, Pos.Index:≥1.00)"
'                        End If
''                    Case "L5165"        'Anti-HBe
''                        If CCur(lsResult) <= 0.9 Then
''                            gArrExamRes(liEquipCode).res = "Positive"
'''                        ElseIf CCur(lsResult) > 1 Then
'''                            gArrExamRes(liEquipCode).res = "Negative"
''                        ElseIf CCur(lsResult) >= 1.49 And CCur(lsResult) < 0.9 Then
''                            gArrExamRes(liEquipCode).res = "Reset"
''                        Else
''                            gArrExamRes(liEquipCode).res = "Negative"
''                        End If
'
''                    Case "L5112"        'Anti-HBs
''                        If CCur(lsResult) <= 12 Then
''                            gArrExamRes(liEquipCode).res = "10" & "(" & lsResult & ")"
''                        ElseIf CCur(lsResult) >= 13 And CCur(lsResult) < 1000 Then
''                            gArrExamRes(liEquipCode).res = lsResult
''                        Else
''                            gArrExamRes(liEquipCode).res = ">1000"
''                        End If
'                    End Select
                    
                    Save_Local_One lRow, liEquipCode, "A", lsGubun
                    
                Else
                    lsExamName = ""
                    
'                    If Not IsNull(mExam) Then
'                        For i = 0 To UBound(mExam, 2)
'                            If Trim(mExam(3, i)) = lsEquipCode Then
'                                lsExamName = Trim(mExam(4, i))
'                                Exit For
'                            End If
'                        Next i
'                    End If
'
'                    If lsEquipCode = "0148" And lsExamName = "" Then
'                        If Not IsNull(mExam) Then
'                            For i = 0 To UBound(mExam, 2)
'                                If Trim(mExam(3, i)) = "0106" Then
'                                    lsExamName = Trim(mExam(4, i))
'                                    Exit For
'                                End If
'                            Next i
'                        End If
'                        If lsExamName <> "" Then lsEquipCode = "0106"
'                    End If
                    
                    liEquipCode = liEquipCode + 1
                    
                    ReDim Preserve gArrExamRes(1 To liEquipCode)
                    
                    gArrExamRes(liEquipCode).EquipCode = lsEquipCode
                    gArrExamRes(liEquipCode).ExamCode = lsEquipCode
                    gArrExamRes(liEquipCode).ExamName = lsExamName
                    gArrExamRes(liEquipCode).SeqNo = ""
                    gArrExamRes(liEquipCode).RefLow = ""
                    gArrExamRes(liEquipCode).RefHigh = ""
                    'gArrExamRes(liEquipCode).RefFlag = ""
                    gArrExamRes(liEquipCode).RefFlag = ""
                    gArrExamRes(liEquipCode).res = lsResult
                    gArrExamRes(liEquipCode).EquipRes = lsResult
                    gArrExamRes(liEquipCode).EquipGubun = Trim(typResult.NoOfAnalyzer)
                    If Trim(typResult.Comment(ii)) = "00" Then
                        gArrExamRes(liEquipCode).Equipcomment = ""
                    Else
                        gArrExamRes(liEquipCode).Equipcomment = typResult.Comment(ii)
                    End If
                    
                    'SetResult1 liEquipCode, i
                    SetLimit liEquipCode
                    
                    Save_Local_One lRow, liEquipCode, "A", lsGubun
                End If
                lResRow = vasRes.DataRowCnt + 1
                If lResRow > vasRes.MaxRows Then
                    vasRes.MaxRows = lResRow
                End If
                vasRes.SetText 1, lResRow, Trim(typResult.SampleID)
                vasRes.SetText 2, lResRow, gArrExamRes(liEquipCode).EquipCode
                vasRes.SetText 3, lResRow, gArrExamRes(liEquipCode).ExamName
                vasRes.SetText 4, lResRow, gArrExamRes(liEquipCode).EquipRes
                vasRes.SetText 5, lResRow, gArrExamRes(liEquipCode).res
                vasRes.SetText 6, lResRow, gArrExamRes(liEquipCode).RefFlag
            End If
    
        
            SetText vasList, "결과", lRow, gResCol
        End If
    Next ii
    
    vasList.Row = lRow
    vasList.Col = 1
    If chkMode.Value = 1 Then
        liRet = 1
    
        liRet = ToServer(lRow, vasList)
        If liRet = 1 Then
            SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
            SetText vasList, "완료", lRow, gResCol
            
            vasList.Row = lRow
            vasList.Col = 1
            vasList.Value = 0
            
            Update_Sample lsID
            DeleteWorkList lsID
            
        Else
            SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
            SetText vasList, "실패", lRow, gResCol
        End If
    End If
    
End Sub

Sub SetResult(ByVal aiRow As Integer, ByVal aiItem As Integer)
    Dim iFloat As Integer
    Dim sTmp As String
    Dim sFormat As String
    
    If Not IsNumeric(gArrExamRes(aiRow).res) Then
        Exit Sub
    End If

'    iFloat = gArrExam(aiItem, 5)
'
'    If iFloat = 0 Then
'        gArrExamRes(aiRow).res = CStr(CCur(gArrExamRes(aiRow).res))
'    Else
'        If IsNumeric(Left(gArrExamRes(aiRow).res, Len(gArrExamRes(aiRow).res) - iFloat)) Then
'            sTmp = CCur(Left(gArrExamRes(aiRow).res, Len(gArrExamRes(aiRow).res) - iFloat))
'        Else
'            sTmp = "0"
'        End If
'        gArrExamRes(aiRow).res = sTmp & "." & Right(gArrExamRes(aiRow).res, iFloat)
'        'If aiItem = 1 Or aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
'        '    gArrExamRes(aiRow).Res = CStr(CCur(Left(gArrExamRes(aiRow).Res, 5 - iFloat)) & "." & Right(gArrExamRes(aiRow).Res, iFloat))
'        'Else
'        '    gArrExamRes(aiRow).Res = CStr(CCur(Left(gArrExamRes(aiRow).Res, 4 - iFloat)) & "." & Right(gArrExamRes(aiRow).Res, iFloat))
'        'End If
'    End If
               
    If IsNumeric(gArrExamRes(aiRow).res) And IsNumeric(gArrExam(aiItem, 6)) Then
        If CCur(gArrExam(aiItem, 6)) > gArrExamRes(aiRow).res Then
            gArrExamRes(aiRow).RefFlag = "L"
        End If
    End If
    If IsNumeric(gArrExamRes(aiRow).res) And IsNumeric(gArrExam(aiItem, 7)) Then
        If CCur(gArrExam(aiItem, 7)) < gArrExamRes(aiRow).res Then
            gArrExamRes(aiRow).RefFlag = "H"
        End If
    End If
    
'    iFloat = gArrExam(aiItem, 8)
'    If IsNumeric(iFloat) Then
'        If CInt(iFloat) = 0 Then
'            sFormat = "#0"
'        ElseIf CInt(iFloat) > 0 Then
'            sFormat = ""
'            sFormat = SetChar(sFormat, CInt(iFloat), 1, "0")
'            sFormat = "0." & sFormat
'        End If
'        If IsNumeric(gArrExamRes(aiRow).res) Then
'            gArrExamRes(aiRow).res = Format(CCur(gArrExamRes(aiRow).res), sFormat)
'        End If
'    End If

End Sub

Sub SetResult1(ByVal aiRow As Integer, ByVal aiItem As Integer)
    Dim iFloat As String
    Dim sTmp As String
    Dim sFormat As String
    Dim sChk As String
    Dim lsGiho, lsRes As String
    Dim i As Integer
    
    If Left(gArrExamRes(aiRow).res, 1) = "<" Or Left(gArrExamRes(aiRow).res, 1) = ">" Then
        lsGiho = Left(gArrExamRes(aiRow).res, 1)
        lsRes = Trim(Mid(gArrExamRes(aiRow).res, 2))
    Else
        lsRes = Trim(gArrExamRes(aiRow).res)
    End If
    
    If Not IsNumeric(lsRes) Then
        Exit Sub
    End If

    iFloat = Trim(GetText(vasTemp, aiItem, 5))
'    If IsNumeric(gArrExamRes(aiRow).res) Then
'        gArrExamRes(aiRow).res = Format(gArrExamRes(aiRow).res, "00000")
'    End If
'    If iFloat = 0 Then
'        gArrExamRes(aiRow).res = CStr(CCur(gArrExamRes(aiRow).res))
'    Else
'        If IsNumeric(Left(gArrExamRes(aiRow).res, Len(gArrExamRes(aiRow).res) - iFloat)) Then
'            sTmp = CCur(Left(gArrExamRes(aiRow).res, Len(gArrExamRes(aiRow).res) - iFloat))
'        Else
'            sTmp = "0"
'        End If
'        gArrExamRes(aiRow).res = sTmp & "." & Right(gArrExamRes(aiRow).res, iFloat)
'        'If aiItem = 1 Or aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
'        '    gArrExamRes(aiRow).Res = CStr(CCur(Left(gArrExamRes(aiRow).Res, 5 - iFloat)) & "." & Right(gArrExamRes(aiRow).Res, iFloat))
'        'Else
'        '    gArrExamRes(aiRow).Res = CStr(CCur(Left(gArrExamRes(aiRow).Res, 4 - iFloat)) & "." & Right(gArrExamRes(aiRow).Res, iFloat))
'        'End If
'    End If
               
    If IsNumeric(lsRes) And IsNumeric(GetText(vasTemp, aiItem, 6)) Then
        If CCur(GetText(vasTemp, aiItem, 6)) > lsRes Then
            gArrExamRes(aiRow).RefFlag = "L"
        End If
    End If
    If IsNumeric(lsRes) And IsNumeric(GetText(vasTemp, aiItem, 7)) Then
        If CCur(GetText(vasTemp, aiItem, 7)) < lsRes Then
            gArrExamRes(aiRow).RefFlag = "H"
        End If
    End If

    iFloat = Trim(GetText(vasTemp, aiItem, 8))

    If IsNumeric(iFloat) Then
        If CInt(iFloat) = 0 Then
            sFormat = "#0"
        ElseIf CInt(iFloat) > 0 Then
            sFormat = ""
            For i = 1 To CInt(iFloat)
                sFormat = sFormat & "0"
            Next i
            
            'sFormat = SetChar(sFormat, CInt(iFloat), 1, "0")
            sFormat = "#0." & sFormat
        End If
        If IsNumeric(lsRes) Then
            gArrExamRes(aiRow).res = lsGiho & Format(CCur(lsRes), sFormat)
        End If
    End If

    
    'CuttOff
    If Trim(GetText(vasTemp, aiItem, 9)) = "1" And IsNumeric(lsRes) = True Then    '크다
        If Trim(GetText(vasTemp, aiItem, 11)) = "1" And Trim(GetText(vasTemp, aiItem, 13)) = "1" Then
            If CCur(lsRes) <= CCur(Trim(GetText(vasTemp, aiItem, 10))) Then
                gArrExamRes(aiRow).res = "Negative"
            ElseIf CCur(lsRes) >= CCur(Trim(GetText(vasTemp, aiItem, 12))) Then
                gArrExamRes(aiRow).res = "Positive"
            Else
                If IsNumeric(Trim(GetText(vasTemp, aiItem, 14))) Then
                    If Trim(GetText(vasTemp, aiItem, 15)) = "1" Then
                        If CCur(lsRes) >= CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    Else
                        If CCur(lsRes) > CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    End If
                Else
                    gArrExamRes(aiRow).res = "Reset"
                End If
            End If
        ElseIf Trim(GetText(vasTemp, aiItem, 11)) = "1" And Trim(GetText(vasTemp, aiItem, 13)) = "0" Then
            If CCur(lsRes) <= CCur(Trim(GetText(vasTemp, aiItem, 10))) Then
                gArrExamRes(aiRow).res = "Negative"
            ElseIf CCur(lsRes) > CCur(Trim(GetText(vasTemp, aiItem, 12))) Then
                gArrExamRes(aiRow).res = "Positive"
            Else
                If IsNumeric(Trim(GetText(vasTemp, aiItem, 14))) Then
                    If Trim(GetText(vasTemp, aiItem, 15)) = "1" Then
                        If CCur(lsRes) >= CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    Else
                        If CCur(lsRes) > CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    End If
                Else
                    gArrExamRes(aiRow).res = "Reset"
                End If
            End If
        ElseIf Trim(GetText(vasTemp, aiItem, 11)) = "0" And Trim(GetText(vasTemp, aiItem, 13)) = "1" Then
            If CCur(lsRes) < CCur(Trim(GetText(vasTemp, aiItem, 10))) Then
                gArrExamRes(aiRow).res = "Negative"
            ElseIf CCur(lsRes) >= CCur(Trim(GetText(vasTemp, aiItem, 12))) Then
                gArrExamRes(aiRow).res = "Positive"
            Else
                If IsNumeric(Trim(GetText(vasTemp, aiItem, 14))) Then
                    If Trim(GetText(vasTemp, aiItem, 15)) = "1" Then
                        If CCur(lsRes) >= CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    Else
                        If CCur(lsRes) > CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    End If
                Else
                    gArrExamRes(aiRow).res = "Reset"
                End If
            End If
        ElseIf Trim(GetText(vasTemp, aiItem, 11)) = "0" And Trim(GetText(vasTemp, aiItem, 13)) = "0" Then
            If CCur(lsRes) < CCur(Trim(GetText(vasTemp, aiItem, 10))) Then
                gArrExamRes(aiRow).res = "Negative"
            ElseIf CCur(lsRes) > CCur(Trim(GetText(vasTemp, aiItem, 12))) Then
                gArrExamRes(aiRow).res = "Positive"
            Else
                If IsNumeric(Trim(GetText(vasTemp, aiItem, 14))) Then
                    If Trim(GetText(vasTemp, aiItem, 15)) = "1" Then
                        If CCur(lsRes) >= CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    Else
                        If CCur(lsRes) > CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    End If
                Else
                    gArrExamRes(aiRow).res = "Reset"
                End If
            End If
        End If
    ElseIf Trim(GetText(vasTemp, aiItem, 9)) = "2" And IsNumeric(lsRes) = True Then     '작다
        If Trim(GetText(vasTemp, aiItem, 11)) = "1" And Trim(GetText(vasTemp, aiItem, 13)) = "1" Then
            If CCur(lsRes) >= CCur(Trim(GetText(vasTemp, aiItem, 10))) Then
                gArrExamRes(aiRow).res = "Negative"
            ElseIf CCur(lsRes) <= CCur(Trim(GetText(vasTemp, aiItem, 12))) Then
                gArrExamRes(aiRow).res = "Positive"
            Else
                If IsNumeric(Trim(GetText(vasTemp, aiItem, 14))) Then
                    If Trim(GetText(vasTemp, aiItem, 15)) = "1" Then
                        If CCur(lsRes) <= CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    Else
                        If CCur(lsRes) < CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    End If
                Else
                    gArrExamRes(aiRow).res = "Reset"
                End If
            End If
        ElseIf Trim(GetText(vasTemp, aiItem, 11)) = "1" And Trim(GetText(vasTemp, aiItem, 13)) = "0" Then
            If CCur(lsRes) >= CCur(Trim(GetText(vasTemp, aiItem, 10))) Then
                gArrExamRes(aiRow).res = "Negative"
            ElseIf CCur(lsRes) < CCur(Trim(GetText(vasTemp, aiItem, 12))) Then
                gArrExamRes(aiRow).res = "Positive"
            Else
                If IsNumeric(Trim(GetText(vasTemp, aiItem, 14))) Then
                    If Trim(GetText(vasTemp, aiItem, 15)) = "1" Then
                        If CCur(lsRes) <= CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    Else
                        If CCur(lsRes) < CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    End If
                Else
                    gArrExamRes(aiRow).res = "Reset"
                End If
            End If
        ElseIf Trim(GetText(vasTemp, aiItem, 11)) = "0" And Trim(GetText(vasTemp, aiItem, 13)) = "1" Then
            If CCur(lsRes) > CCur(Trim(GetText(vasTemp, aiItem, 10))) Then
                gArrExamRes(aiRow).res = "Negative"
            ElseIf CCur(lsRes) <= CCur(Trim(GetText(vasTemp, aiItem, 12))) Then
                gArrExamRes(aiRow).res = "Positive"
            Else
                If IsNumeric(Trim(GetText(vasTemp, aiItem, 14))) Then
                    If Trim(GetText(vasTemp, aiItem, 15)) = "1" Then
                        If CCur(lsRes) <= CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    Else
                        If CCur(lsRes) < CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    End If
                Else
                    gArrExamRes(aiRow).res = "Reset"
                End If
            End If
        ElseIf Trim(GetText(vasTemp, aiItem, 11)) = "0" And Trim(GetText(vasTemp, aiItem, 13)) = "0" Then
            If CCur(lsRes) > CCur(Trim(GetText(vasTemp, aiItem, 10))) Then
                gArrExamRes(aiRow).res = "Negative"
            ElseIf CCur(lsRes) < CCur(Trim(GetText(vasTemp, aiItem, 12))) Then
                gArrExamRes(aiRow).res = "Positive"
            Else
                If IsNumeric(Trim(GetText(vasTemp, aiItem, 14))) Then
                    If Trim(GetText(vasTemp, aiItem, 15)) = "1" Then
                        If CCur(lsRes) <= CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    Else
                        If CCur(lsRes) < CCur(Trim(GetText(vasTemp, aiItem, 14))) Then
                            gArrExamRes(aiRow).res = "Lowtiter"
                        Else
                            gArrExamRes(aiRow).res = "Reset"
                        End If
                    End If
                Else
                    gArrExamRes(aiRow).res = "Reset"
                End If
            End If
        End If
    End If
    
    '2008.12.23 윤영기 HBsAg 만 결과란에 수치결과 넣기
    If gArrExamRes(aiRow).ExamCode = "0024" Then
        gArrExamRes(aiRow).res = gArrExamRes(aiRow).res & " (" & Trim(lsGiho & lsRes) & ")"
    End If
    
    '2009.06.25 윤영기 HBsAb 만 결과란에 수치결과 넣기
    If gArrExamRes(aiRow).ExamCode = "0027" Then
        gArrExamRes(aiRow).res = gArrExamRes(aiRow).res & " (" & Trim(lsGiho & lsRes) & ")"
    End If

End Sub

Sub SetLimit(ByVal aiRow As Integer)
    Dim lsGiho As String
    Dim lsRes As String
    
    lsRes = gArrExamRes(aiRow).res
    If Left(lsRes, 1) = "<" Or Left(lsRes, 1) = ">" Then
        lsGiho = Left(lsRes, 1)
        lsRes = Mid(lsRes, 2)
    Else
        lsRes = Trim(gArrExamRes(aiRow).res)
    End If
    
    If Not IsNumeric(lsRes) Then
        Exit Sub
    End If
    
    Select Case gArrExamRes(aiRow).ExamCode
    Case "0945"     'B-HCG
        If lsGiho = "<" Or CCur(lsRes) < 1.2 Then
            gArrExamRes(aiRow).res = "<1.2"
        End If
    Case "0943"     'SCC
        If lsGiho = "<" Or CCur(lsRes) < 0.1 Then
            gArrExamRes(aiRow).res = "<0.1"
        End If
    Case "0941"     'CA 125
        If lsGiho = "<" Or CCur(lsRes) < 0.1 Then
            gArrExamRes(aiRow).res = "0.1"
        End If
    Case "0942"     'CA 15-3
        If lsGiho = "<" Or CCur(lsRes) < 0.1 Then
            gArrExamRes(aiRow).res = "0.1"
        End If
    Case "0690"     'Estradiol
        If lsGiho = "<" Or CCur(lsRes) < 5 Then
            gArrExamRes(aiRow).res = "<5.0"
        End If
    Case "0648"     'TSI
        If lsGiho = "<" Or CCur(lsRes) < 0.3 Then
            gArrExamRes(aiRow).res = "<0.3"
        End If
    Case "0186"     'Anti-Tg
        If lsGiho = "<" Or CCur(lsRes) < 10 Then
            gArrExamRes(aiRow).res = "<10"
        End If
    Case "0277"     'Anti-Tpo
        If lsGiho = "<" Or CCur(lsRes) < 5 Then
            gArrExamRes(aiRow).res = "<5.0"
        End If
    Case "0040"     'CRP
        If lsGiho = "<" Or CCur(lsRes) < 0.01 Then
            gArrExamRes(aiRow).res = "0.01"
        End If
    Case "0043"     'RA
        If lsGiho = "<" Or CCur(lsRes) < 7 Then
            gArrExamRes(aiRow).res = "<7.00"
        End If
    Case "1057"     'Lp(a)
        If lsGiho = "<" Or CCur(lsRes) < 3 Then
            gArrExamRes(aiRow).res = "<3.0"
        End If
    Case "0170"     'Hapto
        If lsGiho = "<" Or CCur(lsRes) < 3 Then
            gArrExamRes(aiRow).res = "<3.0"
        End If
    Case "0944"     'Total PSA
        If lsGiho = "<" Or CCur(lsRes) < 0.01 Then
            gArrExamRes(aiRow).res = "0.01"
        End If
    Case "1819"     'RPR(Syphilis reagin Test)
        If CCur(lsRes) < 1 Then
            gArrExamRes(aiRow).res = "Non Reactive"
        Else
            gArrExamRes(aiRow).res = "Reactive"
        End If
    End Select
End Sub

Function Save_Local_One(ByVal asRow As Long, ByVal aiIndex As Integer, asSend As String, Optional asGubun As String = "P")
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
    sCnt = ""
    SQL = "Delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & gArrExamRes(aiIndex).EquipCode & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, " & _
            "barcode, examtype, receno, " & _
            "pid, pname, pjumin, page, psex, " & _
            "resdate, seqno, diskno, posno, " & _
            "equipcode, examcode, " & _
            "result, equipres, sendflag, examname, " & _
            "refflag,panicflag, deltaflag, unit, refvalue, panicvalue,examuid ) " & vbCrLf & _
          "VALUES ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 5)) & "', '" & asGubun & "', " & _
          "'" & Trim(GetText(vasList, asRow, 3)) & "', '" & Trim(GetText(vasList, asRow, 4)) & "', '', 0, '', " & _
          "'" & sExamDate & "', '" & gArrExamRes(aiIndex).SeqNo & "', '" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "', " & vbCrLf & _
          "'" & gArrExamRes(aiIndex).EquipCode & "', '" & gArrExamRes(aiIndex).ExamCode & "', " & _
          "'" & gArrExamRes(aiIndex).res & "', '" & gArrExamRes(aiIndex).EquipRes & "', '" & asSend & "', '" & gArrExamRes(aiIndex).ExamName & "', " & vbCrLf & _
          "'" & gArrExamRes(aiIndex).RefFlag & "', '', '', '', " & _
          "'" & gArrExamRes(aiIndex).RefLow & " ~ " & gArrExamRes(aiIndex).RefHigh & "', '', '" & gArrExamRes(aiIndex).EquipGubun & "' ) "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function

Function Save_Local_One_1(ByVal asRow As Long, ByVal aiIndex As Integer, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String
    
    sExamDate = GetDateFull
    
    sCnt = ""
    SQL = "Delete FROM pat_res " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & gArrExam(aiIndex, 1) & "'" & vbCrLf & _
          "  AND barcode = '" & Trim(GetText(vasList, asRow, 2)) & "' "
    
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
    SQL = "INSERT INTO pat_res (examdate, equipno, barcode, examtype, receno, pid, " & _
          "pname, pjumin, page, psex, resdate, seqno, diskno, posno, " & _
          "equipcode, examcode, examtype, result, sendflag, examname, " & _
          "refflag,panicflag, deltaflag, unit, refvalue, panicvalue ) " & vbCrLf & _
          "VALUES ('" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
          "'" & Trim(GetText(vasList, asRow, 2)) & "','" & Trim(GetText(vasList, asRow, 5)) & "', '', " & _
          "'" & Trim(GetText(vasList, asRow, 3)) & "', " & vbCrLf & _
          "'" & Trim(GetText(vasList, asRow, 4)) & "', '', " & _
          "0, '', " & _
          "'" & sExamDate & "', '" & gArrExam(aiIndex, 4) & "', '" & Trim(GetText(vasList, asRow, 6)) & "', '" & Trim(GetText(vasList, asRow, 7)) & "', " & vbCrLf & _
          "'" & gArrExam(aiIndex, 1) & "', '" & gArrExam(aiIndex, 2) & "', '', " & _
          "'" & Trim(GetText(vasList, asRow, gResCol + aiIndex)) & "', '" & asSend & "', '" & gArrExam(aiIndex, 3) & "', " & vbCrLf & _
          "'', '', " & _
          "'', '', " & _
          "'', '') "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
    
End Function


Function Update_Sample(ByVal asID As String)
    SQL = "Update pat_res set sendflag = 'B' " & vbCrLf & _
          "WHERE examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & asID & "' "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Function DeleteWorkList(ByVal asID As String)
    SQL = "Delete from WorkList where Barcode ='" & asID & "'"
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Public Function Set_EqpResultsql(ByVal Testcd As String, ByVal EqpRst As String, ByVal ErrDes As String, ByVal SPCID As String, ByVal INS_CODE As String) As Boolean
On Error GoTo errtrap
    
    'Set cmdSQL = New ADODB.Command
    Dim sDate As String
    
    sDate = GetDateFull
    
    DoSleep 5
    
    SaveRes "InterfaceResult_INSERT_sp " & SPCID & ", " & Testcd & ", " & sDate & ", " & Trim(EqpRst) & ", " & Trim(INS_CODE) & ", " & Trim(ErrDes)
    
    With cmdSQL
        .ActiveConnection = cn_Ser
        .CommandType = adCmdStoredProc
        .CommandText = "InterfaceResult_INSERT_sp"
        .Parameters.Append .CreateParameter("retval", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@i_barcodeNumber", adChar, adParamInput, 11, Trim(SPCID))
        .Parameters.Append .CreateParameter("@i_itemCode", adVarChar, adParamInput, 10, Trim(Testcd))
        .Parameters.Append .CreateParameter("@i_transTimestamp", adChar, adParamInput, 19, sDate)
        .Parameters.Append .CreateParameter("@i_itemResultValue", adVarChar, adParamInput, 1000, Trim(EqpRst))
        .Parameters.Append .CreateParameter("@i_instrumentCode", adChar, adParamInput, 2, Trim(INS_CODE))
        .Parameters.Append .CreateParameter("@i_errorDescription", adVarChar, adParamInput, 100, Trim(ErrDes))
        
        .Execute
    End With
    
    If cmdSQL("retval").Value = 2 Then
        Set_EqpResultsql = False
        MsgBox "결과전송 실패", vbInformation, "알림"
        'Set cmdSQL = Nothing
        Exit Function
    End If
    
    Set_EqpResultsql = True
    'Set cmdSQL = Nothing
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing

    Exit Function
    
errtrap:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cmdSQL Is Nothing Then Set cmdSQL = Nothing
    'Err.Raise Err.Number, Err.Description
End Function

Public Sub SaveRes(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) = "" Then
        MkDir App.Path & "\Log"
    End If
    
    Open App.Path & "\Log\" & Format(Date, "yyyymmdd") & ".log" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Private Sub Picture1_Click()
    frmUser.Show 0
End Sub

Private Sub subClear_Click()

    txtBuff = ""
    
    gCurRow = -1
    ReDim gArrExamRes(0)
    'GetExamCode
    
'vsSpread의 내용을 Clear 한다.
    vasList.Row = 1
    vasList.Col = 1
    vasList.Row2 = vasList.MaxRows
    vasList.Col2 = vasList.MaxCols
    vasList.BlockMode = True
    vasList.Action = 3
    vasList.BackColor = RGB(255, 255, 255)
    vasList.ForeColor = RGB(0, 0, 0)
    vasList.BlockMode = False
    
    ClearSpread vasRes
    
    txtID1 = ""
    txtWkNo1 = ""
    txtPID1 = ""
    txtPName1 = ""
    txtEquip1 = ""
    txtRack1 = ""
    txtResDate1 = ""
    
'    vasList.Row = 1
'    vasList.Col = 1
'    vasList.Row2 = vasList.MaxRows
'    vasList.Col2 = 1
'    vasList.BlockMode = True
'    vasList.Value = 0
'    vasList.BlockMode = False
    
End Sub

Private Sub subClose_Click()
    
    Unload Me
End Sub

Private Sub subCodeSet_Click()
    frmCode.Show
    'GetExamCode
End Sub


Private Sub subComSetup_Click()
    frmConfig.Show 1
End Sub


Private Sub subErrorCode_Click()
    frmErrorCode.Show 1
End Sub

Private Sub subQCInfo_Click()
    frmQCInfo.Show 1
End Sub

Private Sub subSend1_Click()
    subSend1.Checked = True
    subSend2.Checked = False
    
    chkMode.Value = 1
    SaveSetting "MEDIMATE", "CLINILOG", "SendMode", "1"
End Sub

Private Sub subSend2_Click()
    subSend1.Checked = False
    subSend2.Checked = True
    
    chkMode.Value = 0
    SaveSetting "MEDIMATE", "CLINILOG", "SendMode", "0"

End Sub


Private Sub Timer1_Timer()
'    dtpExamDate.Value = Format(Date, "yyyy-mm-dd")
'    sspTime.Caption = Format(Time, "hh:nn")
    
'    If IPU1.ConnectFlag Then
        If MSComm1.CTSHolding = True Then
            lblCA.ForeColor = RGB(0, 255, 0)
        Else
            lblCA.ForeColor = RGB(0, 0, 255)
        End If
'    End If
'
'    If IPU2.ConnectFlag Then
'        If MSComm2.CTSHolding = True Then
'            lblIPU2.ForeColor = RGB(0, 255, 0)
'        Else
'            lblIPU2.ForeColor = RGB(0, 0, 255)
'        End If
'    End If
    
    Dim lsTmp As String * 20
    Dim lInt
    lsTmp = ""
    Call GetPrivateProfileString("DATA", "Days", "", lsTmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(lsTmp)
    If IsNumeric(Trim(txtTemp)) Then
        lInt = CLng(Trim(txtTemp))
    Else
        lInt = 30
    End If
    
    SQL = "Delete from pat_res " & vbCrLf & _
          "WHERE examdate < '" & Format(DateAdd("d", 0 - lInt, CDate(dtpExamDate.Value)), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' "
    res = SendQuery(gLocal, SQL)

End Sub

Private Sub txtBarcode_GotFocus()
    SelectFocus txtBarcode
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    Dim iSch As Integer
    
    If KeyCode = vbKeyReturn Then
        iSch = -1
        For lRow = 1 To vasList.DataRowCnt
            If Trim(GetText(vasList, lRow, 2)) = Trim(txtBarcode) Then
                vasActiveCell vasList, lRow, 2
                iSch = 1
                Exit For
            End If
        Next lRow
        If iSch = -1 Then
            SearchSample Trim(txtBarcode)
        End If
    End If

End Sub

Private Sub txtBuff_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CLINILOG
    End If
End Sub

Private Sub vasList_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
'    iRow1 = BlockRow
'    iRow2 = BlockRow2
'    iCol1 = BlockCol
'    iCol2 = BlockCol2
    
    If BlockRow > BlockRow2 Then
        iRow1 = BlockRow2
        iRow2 = BlockRow
    Else
        iRow1 = BlockRow
        iRow2 = BlockRow2
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        iRow1 = 0
        iRow2 = 0
        Select Case Col
        Case 2
            vasSort vasList, Col
        Case 3
            vasSort vasList, Col
        Case 4
            vasSort vasList, Col, 2
        Case 5
            vasSort vasList, Col, 6
        Case 6
            vasSort vasList, Col, 5
        Case 7
            vasSort vasList, Col, 2
        End Select
    ElseIf Row > 0 Then
        iRow1 = Row
        iRow2 = Row
        
        ClearSpread vasRes
        
        txtID1 = Trim(GetText(vasList, Row, 2))
        txtPID1 = Trim(GetText(vasList, Row, 3))
        txtPName1 = Trim(GetText(vasList, Row, 4))
        txtRack1 = Trim(GetText(vasList, Row, 6)) & "-" & Trim(GetText(vasList, Row, 7))
        txtWkNo1 = Trim(GetText(vasList, Row, 5))
        'txtEquip1 = Mid(SetEquip(Trim(GetText(vasList, Row, gResCol + 1))), 4)
        txtEquip1 = Mid(SetEquip(Trim(GetText(vasList, Row, gResCol + 1))), 10)
        
        SQL = "Select resdate from pat_res " & vbCrLf & _
              "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
              "  AND barcode = '" & Trim(txtID1) & "' "
        res = db_select_Text(gLocal, SQL, txtResDate1)
    
        SQL = "Select barcode, examcode, examname, equipres, result, refflag "
        SQL = SQL & " from pat_res " & vbCrLf & _
              "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
              "  AND barcode = '" & Trim(txtID1) & "' "
        res = db_select_Vas(gLocal, SQL, vasRes)
        vasRes.MaxRows = vasRes.DataRowCnt
    End If
    
    
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lRow, lRow1, lCol As Long
    Dim argSpread As vaSpread
    
    Exit Sub
    
    If Row < 1 Or Row > vasList.DataRowCnt Then Exit Sub
    
    If Col > gResCol Then Exit Sub
    
    SelVas = 1
    
    If Row = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf Row = vasList.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If vasList.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    txtID = Trim(GetText(vasList, Row, 2))
    txtPID = Trim(GetText(vasList, Row, 3))
    txtPName = Trim(GetText(vasList, Row, 4))
    txtRack = Trim(GetText(vasList, Row, 6))
    txtTube = Trim(GetText(vasList, Row, 7))
    txtEquip = Trim(GetText(vasList, Row, 5))
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    
    
    'lCol = gResCol
    lRow = 0
    lRow1 = 0
    
    For lCol = gResCol + 1 To gMaxCol - 1
        If Trim(GetText(vasList, Row, lCol)) <> "" Then
            lRow = lRow + 1
            If lRow <= 20 Then
                Set argSpread = vasRes1
            Else
                Set argSpread = vasRes2
            End If
            If lRow = 21 Then lRow1 = 0
            
            lRow1 = lRow1 + 1
            
            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
            SetText argSpread, Trim(GetText(vasList, Row, lCol)), lRow, 3
            SetText argSpread, Trim(GetText(vasList, 0, lCol)), lRow, 2
            
            vasList.Row = Row
            vasList.Col = lCol
            Select Case vasList.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
                SetText argSpread, "▲", lRow, 4
            Case RGB(0, 127, 255)
                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
                SetText argSpread, "▼", lRow, 4
            Case Else
                SetText argSpread, "", lRow, 4
            End Select
        
        End If
    Next lCol
    
'    For lRow = 1 To 20
'        lCol = lCol + 1
'
'        If Trim(GetText(vaslist, Row, lCol)) <> "" Then
'
'            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
'            vasActiveCell vaslist, lRow, lCol
'            SetText argSpread, Trim(GetText(vaslist, Row, lCol)), lRow, 3
'            SetText argSpread, Trim(GetText(vaslist, 0, lCol)), lRow, 2
'
'            vaslist.Row = Row
'            vaslist.Col = lCol
'            Select Case vaslist.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
'                SetText argSpread, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
'                SetText argSpread, "▼", lRow, 4
'            Case Else
'                SetText argSpread, "", lRow, 4
'            End Select
'        End If
'    Next lRow
'
'    For lRow = 1 To 15
'        lCol = lCol + 1
'
'        If Trim(GetText(vaslist, lRow, lCol)) <> "" Then
'
'            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
'            SetText argSpread, Trim(GetText(vaslist, Row, lCol)), lRow, 3
'            SetText argSpread, Trim(GetText(vaslist, 0, lCol)), lRow, 2
'
'            vaslist.Row = Row
'            vaslist.Col = lCol
'            Select Case vaslist.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
'                SetText argSpread, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
'                SetText argSpread, "▼", lRow, 4
'            Case Else
'                SetText argSpread, "", lRow, 4
'            End Select
'        End If
'    Next lRow
    
    Frame1.Visible = True
    
End Sub

Sub GetComSetup()
    Dim db_tmp As String * 20
    Dim lRow As Long
       
    lRow = 0
        
    db_tmp = ""
    Call GetPrivateProfileString("COM", "Port", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.ComPort = Trim(txtTemp)
                                            
    Call GetPrivateProfileString("COM", "Speed", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.Speed = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "Parity", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.Parity = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "DataBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.DataBit = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "StopBit", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.StopBit = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "RTSEnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.RTSEnable = Trim(txtTemp)

    db_tmp = ""
    Call GetPrivateProfileString("COM", "DTREnable", "", db_tmp, 20, App.Path & "\Interface.ini")
    txtTemp = Trim(db_tmp)
    CA_COM.DTREnable = Trim(txtTemp)

End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow, i, j, z As Long
    
    Dim lsID As String
    Dim liRet As Integer
    'Dim lsID As String
    Dim lsResult As String
    Dim lsEquipCode As String
    Dim lsExamCode As String
    Dim mExam
    Dim lsGubun As String
    
    If KeyCode = vbKeyReturn Then
        lRow = vasList.ActiveRow
        
        SQL = "Select barcode, diskno, posno, examtype from pat_res where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' and barcode = '" & Trim(GetText(vasList, lRow, 2)) & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(GetText(vasList, lRow, 2)) Then
            If MsgBox("입력하신 검체 [" & Trim(GetText(vasList, lRow, 2)) & "]는 장비 " & Trim(gReadBuf(3)) & "의 " & Trim(gReadBuf(1)) & " Rack " & Trim(gReadBuf(2)) & " Position 에서 검사한 것입니다 " & vbCrLf & _
                      " " & vbCrLf & _
                      "저장하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbNo Then
                SetText vasList, Trim(GetText(vasList, lRow, gMaxCol)), lRow, 2
                Exit Sub
            End If
        End If
        
        If MsgBox("결과를 전송하시겠습니까? ", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbYes Then
            
            lsID = Trim(GetText(vasList, lRow, 2))
            mExam = Get_OrderBody(lsID)
            
            If Not IsNull(mExam) Then
                Dim lsEquipQC As String
                Dim lsQCID As String
                
                If Left(Trim(mExam(1, LBound(mExam, 2))), 1) = "Q" Then 'QC 샘플
                    lsEquipQC = Trim(GetText(vasList, lRow, gMaxCol))
                    i = InStr(1, lsEquipQC, ":")
                    If i > 0 Then
                        lsEquipQC = Trim(Left(lsEquipQC, i - 1))
                        
                        SQL = "Select qcid from qcinfo where inscode = '" & gInsCode & "' and equipqc = '" & lsEquipQC & "' "
                        res = db_select_Col(gLocal, SQL)
                        lsQCID = Trim(gReadBuf(0))
                        
                        If Trim(lsQCID) <> Trim(mExam(1, LBound(mExam, 2))) Then
                            MsgBox "잘못된 QC 바코드 입니다. 확인 바랍니다", vbInformation, "바코드 확인"
                            vasList.SetText 2, lRow, Trim(GetText(vasList, lRow, gMaxCol))
                            Exit Sub
                        End If
                    End If
                End If
                
                SetText vasList, Trim(mExam(1, LBound(mExam, 2))), lRow, 3  '등록번호
                SetText vasList, Trim(mExam(2, LBound(mExam, 2))), lRow, 4  '환자명

                If IsNumeric(Trim(mExam(6, LBound(mExam, 2)))) Then
                    SetText vasList, Trim(mExam(5, LBound(mExam, 2))) & "-" & Format(CCur(Trim(mExam(6, LBound(mExam, 2)))), "000"), lRow, 5
                Else
                    SetText vasList, Trim(mExam(5, LBound(mExam, 2))) & "-" & Trim(mExam(6, LBound(mExam, 2))), lRow, 5
                End If
                               
                
            Else
                vasList.Row = lRow
                vasList.Col = 1
                vasList.Value = 1
                SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                SetText vasList, "오류", lRow, gResCol
                
                Exit Sub
            End If
            
            SQL = "select receno from pat_res " & vbCrLf & _
                  "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                  "  and barcode = '" & Trim(GetText(vasList, lRow, gMaxCol)) & "'"
            res = db_select_Col(gLocal, SQL)
            lsGubun = Trim(gReadBuf(0))
            If Trim(lsGubun) = "" Then
                lsGubun = "P"
            End If
            
            SQL = "Update pat_res set " & vbCrLf & _
                  "  barcode = '" & lsID & "', " & vbCrLf & _
                  "  pid = '" & Trim(GetText(vasList, lRow, 3)) & "', " & vbCrLf & _
                  "  pname = '" & Trim(GetText(vasList, lRow, 4)) & "', " & vbCrLf & _
                  "  examtype = '" & Trim(GetText(vasList, lRow, 5)) & "' " & vbCrLf & _
                  "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                  "  and barcode = '" & Trim(GetText(vasList, lRow, gMaxCol)) & "'"
            res = SendQuery(gLocal, SQL)
            
            liRet = 1
            
            ReDim gArrExamRes(1 To gMaxCol - 1 - gResCol)
            
            For i = 1 To gMaxCol - 1 - gResCol
                lsResult = Trim(GetText(vasList, lRow, i + gResCol))
                
                If Trim(lsResult) <> "" Then
                    lsEquipCode = Trim(gArrExam(i, 1))
                    
                    ClearSpread vasTemp
                    SQL = "SELECT EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh, RSGubun " & CR & _
                          "  From EquipExam " & CR & _
                          " WHERE Equip = '" & gEquip & "' " & CR & _
                          "   and EquipCode = '" & lsEquipCode & "' " & vbCrLf & _
                          " Order by seqno "
                    res = db_select_Vas(gLocal, SQL, vasTemp)
                    
                    For j = 1 To vasTemp.DataRowCnt
                        z = LBound(mExam, 2)
                        Do While z <= UBound(mExam, 2)
                            If Trim(GetText(vasTemp, j, 2)) = Trim(mExam(3, z)) Then
                                
                                gArrExamRes(i).EquipCode = Trim(GetText(vasTemp, j, 1))
                                gArrExamRes(i).ExamCode = Trim(GetText(vasTemp, j, 2))
                                gArrExamRes(i).ExamName = Trim(GetText(vasTemp, j, 3))
                                gArrExamRes(i).SeqNo = Trim(GetText(vasTemp, j, 4))
                                gArrExamRes(i).RefLow = Trim(GetText(vasTemp, j, 6))
                                gArrExamRes(i).RefHigh = Trim(GetText(vasTemp, j, 7))
                                gArrExamRes(i).RefFlag = ""
                                gArrExamRes(i).res = lsResult
                                'gArrExamRes(i).EquipGubun = Mid(lsTmp, 9, 1)
                                
                                'SetResult1 i, j
                                
                                
                                Save_Local_One lRow, i, "A", lsGubun
    
                                SetText vasList, gArrExamRes(i).res, lRow, i + gResCol
        
                                If gArrExamRes(i).RefFlag = "H" Then
                                    SetForeColor vasList, lRow, lRow, i + gResCol, i + gResCol, 255, 127, 0
                                ElseIf gArrExamRes(i).RefFlag = "L" Then
                                    SetForeColor vasList, lRow, lRow, i + gResCol, i + gResCol, 0, 127, 255
                                Else
                                    SetForeColor vasList, lRow, lRow, i + gResCol, i + gResCol, 0, 0, 0
                                End If
                                
                                If Set_EqpResultsql(gArrExamRes(i).ExamCode, gArrExamRes(i).res, "", lsID, gInsCode) Then
                                Else
                                    liRet = -1
                                End If
                                
                                Exit For
                            End If
                            z = z + 1
                        Loop
                    Next j
                End If
            Next i
            
            If liRet = 1 Then
                SetBackColor vasList, lRow, lRow, 1, 1, 202, 255, 112
                SetText vasList, "전송", lRow, gResCol
                
                vasList.Row = lRow
                vasList.Col = 1
                vasList.Value = 1
                
                Update_Sample Trim(GetText(vasList, lRow, 2))
                'DeleteWorkList Trim(GetText(vaslist, lRow, 2))
                
            Else
                SetBackColor vasList, lRow, lRow, 1, 1, 255, 0, 0
                SetText vasList, "실패", lRow, gResCol
            End If
            
        End If
    End If
End Sub

Private Sub vasList_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Dim i As Long
    
    If iRow1 < 0 And iRow2 < 0 Then
        iRow1 = Row
        iRow2 = Row
    End If
    
    For i = iRow1 To iRow2
        vasList.Row = i
        vasList.Col = 1
        vasList.Value = 1
    Next i
    vasList.BlockMode = False

End Sub

Private Sub vasSch_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lRow, lRow1, lCol As Long
    Dim argSpread As vaSpread
    
    If Row < 1 Or Row > vasSch.DataRowCnt Then Exit Sub
    
    If Col > gResCol Then Exit Sub
    
    SelVas = 2
    
    If Row = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
    ElseIf Row = vasSch.DataRowCnt Then
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
    Else
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
    End If
    
    If vasSch.DataRowCnt = 1 Then
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
    End If
    
    
    txtID = ""
    txtPID = ""
    txtPName = ""
    txtResDate = ""
    txtEquip = ""
    
    ClearSpread vasRes1
    ClearSpread vasRes2
    
    txtID = Trim(GetText(vasSch, Row, 2))
    txtPID = Trim(GetText(vasSch, Row, 3))
    txtPName = Trim(GetText(vasSch, Row, 4))
    txtRack = Trim(GetText(vasSch, Row, 6))
    txtTube = Trim(GetText(vasSch, Row, 7))
    txtEquip = Trim(GetText(vasSch, Row, 5))
    
    SQL = "Select resdate from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND barcode = '" & Trim(txtID) & "' "
    res = db_select_Text(gLocal, SQL, txtResDate)
    
    
    
    'lCol = gResCol
    lRow = 0
    lRow1 = 0
    For lCol = gResCol + 1 To gMaxCol - 1
        If Trim(GetText(vasSch, Row, lCol)) <> "" Then
            lRow = lRow + 1
            If lRow <= 20 Then
                Set argSpread = vasRes1
            Else
                Set argSpread = vasRes2
            End If
            If lRow = 21 Then lRow1 = 0
            
            lRow1 = lRow1 + 1
            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow1, 1
            SetText argSpread, Trim(GetText(vasSch, Row, lCol)), lRow1, 3
            SetText argSpread, Trim(GetText(vasSch, 0, lCol)), lRow1, 2
            
            vasSch.Row = Row
            vasSch.Col = lCol
            Select Case vasSch.ForeColor
            Case RGB(255, 127, 0)
                SetForeColor argSpread, lRow1, lRow1, 4, 4, 255, 127, 0
                SetText argSpread, "▲", lRow1, 4
            Case RGB(0, 127, 255)
                SetForeColor argSpread, lRow1, lRow1, 4, 4, 0, 127, 255
                SetText argSpread, "▼", lRow1, 4
            Case Else
                SetText argSpread, "", lRow1, 4
            End Select
        
        End If
    Next lCol
    
'    For lRow = 1 To 20
'        lCol = lCol + 1
'
'        If Trim(GetText(vasSch, Row, lCol)) <> "" Then
'
'            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
'            vasActiveCell vasSch, lRow, lCol
'            SetText argSpread, Trim(GetText(vasSch, Row, lCol)), lRow, 3
'            SetText argSpread, Trim(GetText(vasSch, 0, lCol)), lRow, 2
'
'            vasSch.Row = Row
'            vasSch.Col = lCol
'            Select Case vasSch.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
'                SetText argSpread, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
'                SetText argSpread, "▼", lRow, 4
'            Case Else
'                SetText argSpread, "", lRow, 4
'            End Select
'        End If
'    Next lRow
'
'    For lRow = 1 To 15
'        lCol = lCol + 1
'
'        If Trim(GetText(vasSch, lRow, lCol)) <> "" Then
'
'            SetText argSpread, gArrExam(lCol - gResCol, 1), lRow, 1
'            SetText argSpread, Trim(GetText(vasSch, Row, lCol)), lRow, 3
'            SetText argSpread, Trim(GetText(vasSch, 0, lCol)), lRow, 2
'
'            vasSch.Row = Row
'            vasSch.Col = lCol
'            Select Case vasSch.ForeColor
'            Case RGB(255, 127, 0)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 255, 127, 0
'                SetText argSpread, "▲", lRow, 4
'            Case RGB(0, 127, 255)
'                SetForeColor argSpread, lRow, lRow, 4, 4, 0, 127, 255
'                SetText argSpread, "▼", lRow, 4
'            Case Else
'                SetText argSpread, "", lRow, 4
'            End Select
'        End If
'    Next lRow
    
    Frame1.Visible = True

End Sub

Private Sub vasSch_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow, i, j, z As Long
    
    Dim lsID As String
    Dim liRet As Integer
    'Dim lsID As String
    Dim lsResult As String
    Dim lsEquipCode As String
    Dim lsExamCode As String
    Dim mExam
    Dim lsGubun As String
    
    
    If KeyCode = vbKeyReturn Then
        lRow = vasSch.ActiveRow
        
        SQL = "Select barcode, diskno, posno, examtype from pat_res where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' and barcode = '" & Trim(GetText(vasSch, lRow, 2)) & "' "
        res = db_select_Col(gLocal, SQL)
        If Trim(gReadBuf(0)) = Trim(GetText(vasSch, lRow, 2)) Then
            If MsgBox("입력하신 검체 [" & Trim(GetText(vasSch, lRow, 2)) & "]는 장비 " & Trim(gReadBuf(3)) & "의 " & Trim(gReadBuf(1)) & " Rack " & Trim(gReadBuf(2)) & " Position 에서 검사한 것입니다 " & vbCrLf & _
                      " " & vbCrLf & _
                      "저장하시겠습니까?", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbNo Then
                SetText vasSch, Trim(GetText(vasSch, lRow, gMaxCol)), lRow, 2
                Exit Sub
            End If
        End If
        
        If MsgBox("결과를 전송하시겠습니까? ", vbCritical + vbYesNo + vbDefaultButton2, "알림") = vbYes Then
            
            lsID = Trim(GetText(vasSch, lRow, 2))
            mExam = Get_OrderBody(lsID)
            
            If Not IsNull(mExam) Then
                Dim lsEquipQC As String
                Dim lsQCID As String
                
                If Left(Trim(mExam(1, LBound(mExam, 2))), 1) = "Q" Then 'QC 샘플
                    lsEquipQC = Trim(GetText(vasSch, lRow, gMaxCol))
                    i = InStr(1, lsEquipQC, ":")
                    If i > 0 Then
                        lsEquipQC = Trim(Left(lsEquipQC, i - 1))
                        
                        SQL = "Select qcid from qcinfo where inscode = '" & gInsCode & "' and equipqc = '" & lsEquipQC & "' "
                        res = db_select_Col(gLocal, SQL)
                        lsQCID = Trim(gReadBuf(0))
                        
                        If Trim(lsQCID) <> Trim(mExam(1, LBound(mExam, 2))) Then
                            MsgBox "잘못된 QC 바코드 입니다. 확인 바랍니다", vbInformation, "바코드 확인"
                            vasSch.SetText 2, lRow, Trim(GetText(vasSch, lRow, gMaxCol))
                            Exit Sub
                        End If
                    End If
                End If
                
                SetText vasSch, Trim(mExam(1, LBound(mExam, 2))), lRow, 3  '등록번호
                SetText vasSch, Trim(mExam(2, LBound(mExam, 2))), lRow, 4  '환자명

                If IsNumeric(Trim(mExam(6, LBound(mExam, 2)))) Then
                    SetText vasSch, Trim(mExam(5, LBound(mExam, 2))) & "-" & Format(CCur(Trim(mExam(6, LBound(mExam, 2)))), "000"), lRow, 5
                Else
                    SetText vasSch, Trim(mExam(5, LBound(mExam, 2))) & "-" & Trim(mExam(6, LBound(mExam, 2))), lRow, 5
                End If
                               
                
            Else
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 1
                SetBackColor vasSch, lRow, lRow, 1, 1, 255, 0, 0
                SetText vasSch, "오류", lRow, gResCol
                
                Exit Sub
            End If
            
            SQL = "select receno from pat_res " & vbCrLf & _
                  "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                  "  and barcode = '" & Trim(GetText(vasSch, lRow, gMaxCol)) & "'"
            res = db_select_Col(gLocal, SQL)
            lsGubun = Trim(gReadBuf(0))
            If Trim(lsGubun) = "" Then
                lsGubun = "P"
            End If
            
            SQL = "Update pat_res set " & vbCrLf & _
                  "  barcode = '" & lsID & "', " & vbCrLf & _
                  "  pid = '" & Trim(GetText(vasSch, lRow, 3)) & "', " & vbCrLf & _
                  "  pname = '" & Trim(GetText(vasSch, lRow, 4)) & "', " & vbCrLf & _
                  "  examtype = '" & Trim(GetText(vasSch, lRow, 5)) & "' " & vbCrLf & _
                  "where examdate = '" & Format(CDate(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
                  "  and barcode = '" & Trim(GetText(vasSch, lRow, gMaxCol)) & "'"
            res = SendQuery(gLocal, SQL)
            
            liRet = 1
            
            ReDim gArrExamRes(1 To gMaxCol - 1 - gResCol)
            
            For i = 1 To gMaxCol - 1 - gResCol
                lsResult = Trim(GetText(vasSch, lRow, i + gResCol))
                
                If Trim(lsResult) <> "" Then
                    lsEquipCode = Trim(gArrExam(i, 1))
                    
                    ClearSpread vasTemp
                    SQL = "SELECT EquipCode,ExamCode, ExamName, Seqno, PointSize, RefLow, RefHigh, RSGubun " & CR & _
                          "  From EquipExam " & CR & _
                          " WHERE Equip = '" & gEquip & "' " & CR & _
                          "   and EquipCode = '" & lsEquipCode & "' " & vbCrLf & _
                          " Order by seqno "
                    res = db_select_Vas(gLocal, SQL, vasTemp)
                    
                    For j = 1 To vasTemp.DataRowCnt
                        z = LBound(mExam, 2)
                        Do While z <= UBound(mExam, 2)
                            If Trim(GetText(vasTemp, j, 2)) = Trim(mExam(3, z)) Then
                                
                                gArrExamRes(i).EquipCode = Trim(GetText(vasTemp, j, 1))
                                gArrExamRes(i).ExamCode = Trim(GetText(vasTemp, j, 2))
                                gArrExamRes(i).ExamName = Trim(GetText(vasTemp, j, 3))
                                gArrExamRes(i).SeqNo = Trim(GetText(vasTemp, j, 4))
                                gArrExamRes(i).RefLow = Trim(GetText(vasTemp, j, 6))
                                gArrExamRes(i).RefHigh = Trim(GetText(vasTemp, j, 7))
                                gArrExamRes(i).RefFlag = ""
                                gArrExamRes(i).res = lsResult
                                'gArrExamRes(i).EquipGubun = Mid(lsTmp, 9, 1)
                                
                                'SetResult1 i, j
                                
                                
                                Save_Local_One lRow, i, "A", lsGubun
    
                                SetText vasList, gArrExamRes(i).res, lRow, i + gResCol
        
                                If gArrExamRes(i).RefFlag = "H" Then
                                    SetForeColor vasList, lRow, lRow, i + gResCol, i + gResCol, 255, 127, 0
                                ElseIf gArrExamRes(i).RefFlag = "L" Then
                                    SetForeColor vasList, lRow, lRow, i + gResCol, i + gResCol, 0, 127, 255
                                Else
                                    SetForeColor vasList, lRow, lRow, i + gResCol, i + gResCol, 0, 0, 0
                                End If
                                
                                If Set_EqpResultsql(gArrExamRes(i).ExamCode, gArrExamRes(i).res, "", lsID, gInsCode) Then
                                Else
                                    liRet = -1
                                End If
                                
                                Exit For
                            End If
                            z = z + 1
                        Loop
                    Next j
                End If
            Next i
            
            If liRet = 1 Then
                SetBackColor vasSch, lRow, lRow, 1, 1, 202, 255, 112
                SetText vasSch, "전송", lRow, gResCol
                
                vasSch.Row = lRow
                vasSch.Col = 1
                vasSch.Value = 1
                
                Update_Sample Trim(GetText(vasSch, lRow, 2))
                'DeleteWorkList Trim(GetText(vasSch, lRow, 2))
                
            Else
                SetBackColor vasSch, lRow, lRow, 1, 1, 255, 0, 0
                SetText vasSch, "실패", lRow, gResCol
            End If
            
        End If
    End If

End Sub

Function SetEquip(asEquip As String) As String
'    Select Case asEquip
'    Case "41"
'        SetEquip = "06 TBA1"
'    Case "42"
'        SetEquip = "07 TBA2"
'    Case "43"
'        SetEquip = "08 TBA3"
'    Case "51"
'        SetEquip = "10 Architect"
'    Case "61"
'        SetEquip = "09 Modular"
'    Case "10"
'        SetEquip = "10 Architect"
'    Case "09"
'        SetEquip = "09 Modular"
'    Case "06"
'        SetEquip = "06 TBA1"
'    Case "07"
'        SetEquip = "07 TBA2"
'    Case "08"
'        SetEquip = "08 TBA3"
'    Case "10"
'        SetEquip = "10 Architect"

    '국립암센터
    Select Case asEquip
    Case "41"
        SetEquip = "200FR 3  200FR1"
    Case "42"
        SetEquip = "200FR 2  200FR2"
    Case "43"
        SetEquip = "200FR 1  200FR3"
    Case "51"
        SetEquip = "Arch 3   Architect1"
    Case "52"
        SetEquip = "Arch 2   Architect2"
    Case "53"
        SetEquip = "Arch 1   Architect3"
    Case "61"
        SetEquip = "Centaur2 Centaur2"
    Case "62"
        SetEquip = "Centaur1 Centaur1"
    End Select
End Function

Function SetErrDesc(asErr As String) As String

    Dim sData As String
    
    sData = Trim(asErr)
    
    SQL = "Select tlaerror, chgerror, errordesc from errorcode where tlaerror = '" & Trim(asErr) & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) = Trim(asErr) Then
        sData = Trim(gReadBuf(1))
    End If
    
    SetErrDesc = sData
    
End Function

Function ToServer(ByVal argSpcRow As Integer, argSpread As vaSpread, Optional asFlag As Integer = 0) As Integer
    Dim i           As Integer
    Dim j           As Integer
    
    Dim lsID        As String
    
    Dim lRow        As Long
    Dim lsMsg       As String
    Dim lsEqFlag    As String
    
    Dim sRet        As String
    
    Dim sParam As String
    
    Dim sResRow As Long
    
    Dim sEquip  As String
    Dim sEquip1 As String
    
    ToServer = -1
    
    lRow = argSpcRow
    
    If lRow < 1 Or lRow > argSpread.DataRowCnt Then Exit Function
    
    lsID = Trim(GetText(argSpread, lRow, 2))
    
    If lsID = "" Then Exit Function
    
    If IsNumeric(lsID) = False Or Len(lsID) < 11 Then Exit Function
    
    ClearSpread vasTemp

    SQL = "Select equipcode, examcode, examname, result, result, pid, pname, receno " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "where  examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & lsID & "' " & vbCrLf & _
          "  and result <> '' "
    
    '2010.04.02 이상은 추가
    SQL = SQL & CR & " And sendflag <> 'C' "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
'    lsMsg = ""
'    lsEqFlag = ""
'''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'    SQL = " Select message From pat_resmemo " & vbCrLf & _
'          " Where examdate = '" & Format(dtpExamDate.Value, "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & Trim(lsID) & "' "
'    res = db_select_Col(gLocal, SQL)
'    If res > 0 Then
'        lsMsg = "XE2100A : " & Trim(gReadBuf(0))
'    End If
'''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    'vasTemp.Visible = True
    If vasTemp.DataRowCnt < 1 Then Exit Function
      
    'Save_Raw_Data lsID & " : 서버 결과 전송 시작"
    'Save_Raw_Data lsID & " : 장부 정보 가져오기"

    On Error GoTo ErrHandle
    
    sParam = ""
    
    For sResRow = 1 To vasTemp.DataRowCnt
        sEquip = ""
        SQL = " Select examuid From pat_res "
        SQL = SQL & CR & " Where examdate = '" & Format(Trim(dtpExamDate.Value), "yyyymmdd") & "' "
        SQL = SQL & CR & " And barcode = '" & lsID & "' "
        SQL = SQL & CR & " And examcode = '" & Trim(GetText(vasTemp, sResRow, 2)) & "' "
        res = db_select_Col(gLocal, SQL)
        sEquip = Trim(gReadBuf(0))
        
        sEquip1 = ""
        sEquip1 = SetEquip(sEquip)
        sEquip = Trim(Mid(sEquip1, 1, 9))
        
        If Trim(GetText(vasTemp, sResRow, 2)) <> "" Then
            sParam = sParam & "<Table>" & _
                    "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                    "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                    "<USERID><![CDATA[" & gServerID & "]]></USERID>" & _
                    "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                    "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                    "<P0><![CDATA[" & lsID & "]]></P0>" & _
                    "<P1><![CDATA[" & Trim(GetText(vasTemp, sResRow, 2)) & "]]></P1>" & _
                    "<P2><![CDATA[" & Trim(GetText(vasTemp, sResRow, 5)) & "]]></P2>" & _
                    "<P3><![CDATA[]]></P3>" & _
                    "<P4><![CDATA[" & sEquip & "]]></P4>" & _
                    "<P5><![CDATA[" & gIFUser & "]]></P5>" & _
                    "<P6><![CDATA[]]></P6>" & _
                    "<P7><![CDATA[" & lsMsg & "]]></P7>" & _
                    "<P8><![CDATA[]]></P8>" & _
                    "<P9><![CDATA[]]></P9>" & _
                    "</Table>"
            SQL = "Update pat_res set sendflag = 'C' " & vbCrLf & _
                  "where equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and barcode = '" & lsID & "' and examcode = '" & Trim(GetText(vasTemp, sResRow, 2)) & "'"
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
                Exit Function
            End If
        End If
    Next
    
    If Trim(sParam) <> "" Then
        sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
        
        Online_Result_Qry sParam
        
        ToServer = 1
    End If
    
    'Save_Raw_Data lsID & " : 서버 결과 전송 완료!"

    Exit Function

ErrHandle:
    'Save_Raw_Data Err.Number & " : " & Err.Description & vbCrLf & _
                  SQL
    Resume Next
End Function


VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmInterface 
   Caption         =   " (1번) COBAS H 232"
   ClientHeight    =   10440
   ClientLeft      =   1845
   ClientTop       =   315
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
   ScaleHeight     =   11261.12
   ScaleMode       =   0  '사용자
   ScaleWidth      =   30990.76
   Begin FPSpread.vaSpread argSpread 
      Height          =   690
      Left            =   1845
      TabIndex        =   74
      Top             =   6435
      Visible         =   0   'False
      Width           =   5730
      _Version        =   393216
      _ExtentX        =   10107
      _ExtentY        =   1217
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
      SpreadDesigner  =   "frmInterface.frx":030A
   End
   Begin VB.TextBox txtTemp 
      Height          =   375
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   2475
      Left            =   2010
      TabIndex        =   72
      Top             =   5910
      Visible         =   0   'False
      Width           =   6465
   End
   Begin VB.TextBox txtRece 
      Height          =   1575
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   71
      Top             =   7650
      Visible         =   0   'False
      Width           =   11745
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   555
      Left            =   8640
      TabIndex        =   70
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Threed.SSPanel spErr 
      Height          =   495
      Left            =   10350
      TabIndex        =   65
      Top             =   1200
      Visible         =   0   'False
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   873
      _Version        =   131072
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCol1 
      Caption         =   "<"
      Height          =   525
      Left            =   2880
      TabIndex        =   64
      Top             =   1860
      Width           =   285
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   1508
      _Version        =   131072
      ForeColor       =   4194304
      BackColor       =   16056319
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " (1번)  COBAS H 232   INTERFACE"
      BevelOuter      =   0
      Alignment       =   1
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   4440
         Top             =   330
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   3990
         Top             =   300
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmErr 
         Enabled         =   0   'False
         Left            =   2940
         Top             =   300
      End
      Begin MSComCtl2.DTPicker dtpToday 
         Height          =   345
         Left            =   7020
         TabIndex        =   48
         Top             =   480
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   96993281
         CurrentDate     =   40534
      End
      Begin VB.TextBox Text_Today 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   330
         TabIndex        =   4
         Text            =   "2002/02/18"
         Top             =   -90
         Visible         =   0   'False
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtpToday_1 
         Height          =   345
         Left            =   7020
         TabIndex        =   63
         Top             =   60
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   96993281
         CurrentDate     =   40534
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   5850
         TabIndex        =   5
         Top             =   120
         Width           =   1020
      End
   End
   Begin VB.TextBox txtBuff2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8070
      MultiLine       =   -1  'True
      TabIndex        =   61
      Top             =   6930
      Visible         =   0   'False
      Width           =   6930
   End
   Begin VB.TextBox txtBuff 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   1230
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   6690
      Visible         =   0   'False
      Width           =   6930
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   735
      Left            =   11910
      TabIndex        =   50
      Top             =   2850
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1296
      _Version        =   131072
      ForeColor       =   16777215
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "이전"
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   855
      Left            =   11910
      TabIndex        =   51
      Top             =   1800
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1508
      _Version        =   131072
      ForeColor       =   16777215
      BackColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "미등록건수"
   End
   Begin Threed.SSPanel spMissResPast 
      Height          =   2475
      Left            =   11910
      TabIndex        =   52
      Top             =   3600
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   4366
      _Version        =   131072
      ForeColor       =   16777215
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   48
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0"
   End
   Begin FPSpread.vaSpread vasTransChk 
      Height          =   2265
      Left            =   6360
      TabIndex        =   41
      Top             =   5340
      Visible         =   0   'False
      Width           =   5655
      _Version        =   393216
      _ExtentX        =   9975
      _ExtentY        =   3995
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
      SpreadDesigner  =   "frmInterface.frx":0579
   End
   Begin FPSpread.vaSpread vas_print 
      Height          =   3015
      Left            =   2940
      TabIndex        =   40
      Top             =   3480
      Visible         =   0   'False
      Width           =   10875
      _Version        =   393216
      _ExtentX        =   19182
      _ExtentY        =   5318
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
      MaxCols         =   8
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":07E8
   End
   Begin FPSpread.vaSpread vasOrder 
      Height          =   2835
      Left            =   1110
      TabIndex        =   38
      Top             =   5880
      Visible         =   0   'False
      Width           =   7275
      _Version        =   393216
      _ExtentX        =   12832
      _ExtentY        =   5001
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
      MaxCols         =   3
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":2183
   End
   Begin VB.CommandButton Command4 
      Caption         =   "test"
      Height          =   615
      Left            =   11160
      TabIndex        =   37
      Top             =   9570
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton cmdWorkList 
      Caption         =   "WorkList"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "새굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10710
      Picture         =   "frmInterface.frx":39F8
      Style           =   1  '그래픽
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   585
      Left            =   11280
      TabIndex        =   33
      Top             =   8850
      Visible         =   0   'False
      Width           =   1365
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
      Height          =   1455
      Left            =   390
      TabIndex        =   6
      Top             =   7800
      Visible         =   0   'False
      Width           =   13740
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "전송"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   10170
      TabIndex        =   27
      Top             =   150
      Width           =   1185
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
      ForeColor       =   &H00008000&
      Height          =   705
      Left            =   8940
      Style           =   1  '그래픽
      TabIndex        =   24
      Top             =   150
      Value           =   1  '확인
      Width           =   1185
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check1"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1080
      TabIndex        =   23
      Top             =   2010
      Width           =   195
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   255
      Left            =   420
      TabIndex        =   22
      Top             =   2010
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   450
      _Version        =   131072
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "번호"
      BevelOuter      =   0
   End
   Begin FPSpread.vaSpread vasExam 
      Height          =   7605
      Left            =   360
      TabIndex        =   32
      Top             =   1830
      Width           =   11475
      _Version        =   393216
      _ExtentX        =   20241
      _ExtentY        =   13414
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   0
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   16
      MaxRows         =   30
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmInterface.frx":42C2
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   705
      Left            =   6120
      TabIndex        =   30
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin FPSpread.vaSpread vasResTemp 
      Height          =   2775
      Left            =   3630
      TabIndex        =   25
      Top             =   4740
      Visible         =   0   'False
      Width           =   3375
      _Version        =   393216
      _ExtentX        =   5953
      _ExtentY        =   4895
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
      SpreadDesigner  =   "frmInterface.frx":4FDA
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   2805
      Left            =   1440
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
      _Version        =   393216
      _ExtentX        =   3625
      _ExtentY        =   4948
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
      SpreadDesigner  =   "frmInterface.frx":5249
   End
   Begin FPSpread.vaSpread vasID 
      Height          =   7335
      Left            =   450
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   8115
      _Version        =   393216
      _ExtentX        =   14314
      _ExtentY        =   12938
      _StockProps     =   64
      ColHeaderDisplay=   0
      ColsFrozen      =   1
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   17
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmInterface.frx":9744
      UserResize      =   2
   End
   Begin FPSpread.vaSpread vasTemp1 
      Height          =   1125
      Left            =   9990
      TabIndex        =   20
      Top             =   2760
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
      SpreadDesigner  =   "frmInterface.frx":F866
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   285
      Left            =   450
      TabIndex        =   19
      Top             =   1950
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   503
      _Version        =   131072
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "번호"
      BevelOuter      =   0
   End
   Begin FPSpread.vaSpread vasTemp 
      Height          =   1125
      Left            =   9570
      TabIndex        =   18
      Top             =   2670
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
      SpreadDesigner  =   "frmInterface.frx":FAD5
   End
   Begin FPSpread.vaSpread vasRes 
      Height          =   6975
      Left            =   5430
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   6195
      _Version        =   393216
      _ExtentX        =   10927
      _ExtentY        =   12303
      _StockProps     =   64
      ColHeaderDisplay=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   16777215
      MaxCols         =   8
      ScrollBars      =   2
      SpreadDesigner  =   "frmInterface.frx":FD44
   End
   Begin VB.CommandButton cmdCall 
      Caption         =   "Local Data 불러오기"
      CausesValidation=   0   'False
      Height          =   435
      Left            =   1935
      TabIndex        =   14
      Top             =   1275
      Width           =   2295
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "화면 초기화"
      CausesValidation=   0   'False
      Height          =   435
      Left            =   375
      TabIndex        =   13
      Top             =   1275
      Width           =   1515
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   10065
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
            TextSave        =   "2012-03-05"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 4:59"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "메디메이트 ☎(02)6205-1751"
            TextSave        =   "메디메이트 ☎(02)6205-1751"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   585
      Left            =   3870
      TabIndex        =   9
      Top             =   5250
      Width           =   1665
   End
   Begin VB.CommandButton Command_Config 
      Caption         =   "통신설정"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   11370
      TabIndex        =   8
      Top             =   150
      Width           =   1185
   End
   Begin VB.CommandButton Command_close 
      Caption         =   "종료"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   13770
      TabIndex        =   1
      Top             =   150
      Width           =   1185
   End
   Begin VB.CommandButton Command_setup 
      Caption         =   "코드설정"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   12570
      TabIndex        =   0
      Top             =   150
      Width           =   1185
   End
   Begin VB.TextBox txtErr 
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   10920
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   2
      Top             =   7560
      Visible         =   0   'False
      Width           =   1245
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
      Height          =   315
      Left            =   8940
      TabIndex        =   7
      Top             =   7830
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Frame Frame7 
      Height          =   9075
      Left            =   120
      TabIndex        =   11
      Top             =   990
      Width           =   14925
      Begin VB.CommandButton cmdUser 
         Caption         =   "사용자관리"
         Height          =   465
         Left            =   13170
         TabIndex        =   66
         Top             =   8490
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdResInsert 
         Caption         =   "결과수기입력"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   8250
         TabIndex        =   49
         Top             =   240
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   210
         ScaleHeight     =   435
         ScaleWidth      =   14595
         TabIndex        =   46
         Top             =   8520
         Width           =   14625
         Begin VB.Label Label7 
            BackStyle       =   0  '투명
            Caption         =   "[장비2]"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   90
            TabIndex        =   69
            Top             =   390
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblConnect2 
            BackStyle       =   0  '투명
            Caption         =   "연결 대기중."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1050
            TabIndex        =   68
            Top             =   390
            Visible         =   0   'False
            Width           =   11475
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   "[장비1]"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   90
            TabIndex        =   67
            Top             =   120
            Width           =   915
         End
         Begin VB.Label lblConnect1 
            BackStyle       =   0  '투명
            Caption         =   "연결 대기중."
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1050
            TabIndex        =   47
            Top             =   120
            Width           =   11475
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   9780
         Picture         =   "frmInterface.frx":13B9A
         ScaleHeight     =   285
         ScaleWidth      =   315
         TabIndex        =   43
         Top             =   8760
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13170
         TabIndex        =   42
         Top             =   8310
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "결과 출력"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13050
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton chkStart 
         Caption         =   "시작"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "새굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13260
         TabIndex        =   36
         Top             =   1080
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   12690
         TabIndex        =   26
         Top             =   1200
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CommandButton cmdDown 
         Height          =   525
         Left            =   8460
         Picture         =   "frmInterface.frx":14124
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   8700
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdUp 
         Height          =   525
         Left            =   7290
         Picture         =   "frmInterface.frx":14256
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   8700
         Visible         =   0   'False
         Width           =   705
      End
      Begin Threed.SSPanel SSPanel6 
         Height          =   735
         Left            =   11820
         TabIndex        =   53
         Top             =   5280
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   1296
         _Version        =   131072
         ForeColor       =   16777215
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "오늘"
      End
      Begin Threed.SSPanel spMissResNow 
         Height          =   2475
         Left            =   11820
         TabIndex        =   54
         Top             =   6030
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   4366
         _Version        =   131072
         ForeColor       =   16777215
         BackColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   48
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0"
      End
      Begin VB.Label lblUser 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   13590
         TabIndex        =   45
         Top             =   8880
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "보고자 :"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12120
         TabIndex        =   44
         Top             =   8400
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   5850
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   5310
      Width           =   2175
   End
   Begin VB.CommandButton cmdQC 
      Caption         =   "QC"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   31
      Top             =   300
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdResSave 
      Caption         =   "결과저장"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "새굴림"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5580
      TabIndex        =   28
      Top             =   540
      Visible         =   0   'False
      Width           =   1725
   End
   Begin Threed.SSPanel spNotTrans 
      Height          =   6315
      Left            =   -420
      TabIndex        =   55
      Top             =   8130
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   11139
      _Version        =   131072
      BackColor       =   15724527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FPSpread.vaSpread vasServerTemp 
         Height          =   1215
         Left            =   1740
         TabIndex        =   62
         Top             =   3120
         Visible         =   0   'False
         Width           =   3225
         _Version        =   393216
         _ExtentX        =   5689
         _ExtentY        =   2143
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
         SpreadDesigner  =   "frmInterface.frx":14385
      End
      Begin VB.CommandButton cmdNotExit 
         Caption         =   "종료"
         Height          =   375
         Left            =   5070
         TabIndex        =   60
         Top             =   150
         Width           =   1305
      End
      Begin FPSpread.vaSpread vasNotResult 
         Height          =   5475
         Left            =   210
         TabIndex        =   59
         Top             =   600
         Width           =   6195
         _Version        =   393216
         _ExtentX        =   10927
         _ExtentY        =   9657
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
         MaxCols         =   5
         MaxRows         =   30
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":145F4
      End
      Begin VB.CommandButton cmdNotResult 
         Caption         =   "조회"
         Height          =   375
         Left            =   3270
         TabIndex        =   58
         Top             =   150
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpNotDate 
         Height          =   315
         Left            =   1500
         TabIndex        =   56
         Top             =   180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Format          =   96993281
         CurrentDate     =   40589
      End
      Begin VB.Label Label3 
         Caption         =   "검사일자"
         Height          =   195
         Left            =   540
         TabIndex        =   57
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Menu subUp 
      Caption         =   "검체번호 변경"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================
'2008.03.24 포항선린병원
'================================================

'vasID

Const colCheckBox = 1
Const colBarcode = 2
Const colReceDate = 3
Const colExamDate = 4
Const colExamTime = 5
Const colTransTime = 6
Const colPID = 7
Const colPName = 8
Const colEquipCode = 9
Const colExamCode = 10
Const colReceCode = 11
Const colExamName = 12
Const colEquipRes = 13
Const colResult = 14
Const colState = 15
Const colErrState = 16


'''Const colSeqNo = 3
'''Const colReceno = 4
'''Const colPID = 5
'''Const colPName = 6
'''Const colPSex = 7
'''Const colPAge = 8
'''Const colPJumin = 9
'''Const colRack = 10
'''Const colTube = 11
'''Const colReceDate = 12
'''Const colState = 13
'''
'''Const colOrd = 14
'''Const colRes = 15
'''Const colDate = 16
'''Const colTime = 17
'''
'''Const colResult = 18
Dim colResult1 As Long
Dim gSendData1 As String
Dim gSendData2 As String

Dim gReceData1 As String
Dim gReceData2 As String

Dim gAllData1 As String
Dim gAllData2 As String


'vasRes
'Const colEquipCode = 1
'Const colExamCode = 2
'Const colExamName = 3
'Const colResult = 4
'Const colSeq = 5
'Const colRCheck = 6
'
''2004/10/21 이상은
''Const colRefLow = 7
'Const colResult1 = 7
'Const colRefHigh = 8

Dim gRow As Long

Dim gsBarCode As String
Dim gsPID As String
Dim gsRackNo As String
Dim gsPosNo As String
Dim gsResDateTime As String
Dim gsSeqNo As String
Dim gsExamCode As String
Dim gsExamName As String
Dim gsOrder As String
Dim gsResult As String

Dim PreRack As String
Dim PrePos As String
Dim PreRow As Long

Dim gwTmp1 As String
Dim gwTmp2 As String

'************************************************
Dim in_spc_no$, spc_no$(), tst_cd$(), tst_nm$()
Dim spc_cd$(), tst_frct_cd$(), tst_frct_nm$()
Dim tst_dte$(), tst_time$(), work_no$()
Dim pt_no$(), pt_nm$(), sex$(), birthday$(), intbase$()

Dim acpt_no$()

Dim rv As Integer
Dim vTemp As String
'************************************************

Function SetResult(asResult As String, aiItem As Integer) As String
    Dim iFloat As Integer

    If Not IsNumeric(asResult) Then
        Exit Function
    End If

    Select Case aiItem
'    Case 1, 14, 15, 16, 17, 18, 30
'        iFloat = 2
    Case 8
        iFloat = 0
    Case 1, 2, 14, 15, 16, 17, 18, 24, 30
        iFloat = 2
    Case Else
        iFloat = 1
    End Select

    If iFloat = 0 Then
        SetResult = CStr(CCur(asResult))
    Else
        If aiItem = 1 Then
            SetResult = CStr(Format(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat), "#0.0"))
        ElseIf aiItem = 14 Or aiItem = 15 Or aiItem = 16 Or aiItem = 17 Or aiItem = 18 Then
            SetResult = CStr(CCur(Left(asResult, 5 - iFloat)) & "." & Right(asResult, iFloat))
        Else
            SetResult = CStr(CCur(Left(asResult, 4 - iFloat)) & "." & Right(asResult, iFloat))
        End If
    End If
    
        
'사용안함
'    If aiItem = 1 Then
'        SetResult = Format(SetResult, "#0.0")
'    End If
End Function

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasExam.DataRowCnt
            vasExam.Row = iRow
            vasExam.Col = 1
            
            vasExam.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasExam.DataRowCnt
            vasExam.Row = iRow
            vasExam.Col = 1
            
            vasExam.Value = 0
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

Private Sub cmdCall_Click()
    Dim iRow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim y As Integer
    Dim sResFlag As String
    Dim sRes As String

    Dim sResult As String

    ClearSpread vasExam
    
'''    check , barcode, examtime, rece, receno, pname, equipcode, examcode, rececode, examname, equipres, result, sendflag, errorstate

    SQL = "select barcode, recedate, examdate, examtime, resdate, pid, pname, equipcode, examcode, rececode, examname, equipres, result, sendflag, bigo " & _
          "from pat_res " & _
          "where examdate between '" & Format(dtpToday_1, "yyyymmdd") & "' and '" & Format(dtpToday, "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "group by examdate, examtime, barcode, recedate, resdate, pid, pname, equipcode, examcode, rececode, examname, equipres, result, sendflag, bigo "
    res = db_select_Vas(gLocal, SQL, vasExam, vasExam.DataRowCnt + 1, 2)
    
'''    ClearSpread vaSpread1
'''
'''    res = db_select_Vas(gLocal, SQL, vaSpread1)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
'''    vasSort vasExam, colRack, colTube
    
    For iRow = 1 To vasExam.DataRowCnt
        If Trim(GetText(vasExam, iRow, colErrState)) = "Positive" Then
            SetForeColor vasExam, iRow, iRow, colErrState, colErrState, 255, 0, 0
            
        End If
        
        Select Case Trim(GetText(vasExam, iRow, colState))
        Case "B"
            SetBackColor vasExam, iRow, iRow, colCheckBox, colState, 255, 250, 205
            SetText vasExam, "결과", iRow, colState
        Case "C"
            SetBackColor vasExam, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasExam, "완료", iRow, colState
        Case Else
            SetBackColor vasExam, iRow, iRow, 1, colState, 255, 255, 255
            SetText vasExam, "", iRow, colState
        End Select
    
        '결과 불러오기
        ClearSpread vasTemp
        
'''        SQL = " Select examcode, result, subcode, equipcode From pat_res " & vbCrLf & _
'''              " Where equipno = '" & gEquip & "' " & vbCrLf & _
'''              " And examdate between '" & Format(dtpToday_1, "yyyymmdd") & "' and '" & Format(dtpToday, "yyyymmdd") & "' " & vbCrLf & _
'''              " And barcode = '" & Trim(GetText(vasExam, iRow, colBarcode)) & "' " & vbCrLf & _
'''              " And pid = '" & Trim(GetText(vasExam, iRow, colPID)) & "' "
'''        res = db_select_Vas(gLocal, SQL, vasTemp)
'''
'''        For i = 1 To vasTemp.DataRowCnt
'''            For j = 1 To UBound(gArrEquip)
'''                If Trim(GetText(vasTemp, i, 4)) = gArrEquip(j, 2) Then
'''                    y = 18 + ((gArrEquip(j, 1)) - 1) * 4
'''                    Exit For
'''                End If
'''            Next j
'''
'''            sResult = Trim(GetText(vasTemp, i, 2))
'''            SetText vasExam, sResult, iRow, y
''''            If gArrEquip(j, 2) = "HBV-HPS" Then
''''                sResFlag = ""
''''                sRes = ""
''''                If Mid(sResult, 1, 1) = ">" Or Mid(sResult, 1, 1) = "<" Then
''''                    sResFlag = Mid(sResult, 1, 1)
''''                    sRes = Mid(sResult, 2)
''''                Else
''''                    sRes = sResult
''''                End If
''''                If IsNumeric(sRes) = True Then
''''                    sRes = CCur(sRes) * 5.82
''''                    sRes = Format(sRes, "###,###,###,###")
''''                    If Right(sRes, 1) = "." Then
''''                        sRes = Mid(sRes, 1, Len(sRes) - 1)
''''                    End If
''''                    vasExam.SetText vasExam.MaxCols, iRow, Trim(sResFlag & sRes)
''''                End If
''''            End If
'''        Next i
    Next iRow
    TransCheck
    vasExam.RowHeight(-1) = 25
End Sub

Private Sub cmdCol1_Click()
    If cmdCol1.Caption = "<" Then
        cmdCol1.Caption = ">"
        vasExam.Col = 4
        vasExam.ColHidden = True
        vasExam.Col = 5
        vasExam.ColHidden = True
        vasExam.Col = 6
        vasExam.ColHidden = True
        vasExam.Col = 7
        vasExam.ColHidden = True
        vasExam.Col = 11
        vasExam.ColHidden = True
        vasExam.Col = 13
        vasExam.ColHidden = True

    Else
        cmdCol1.Caption = "<"
        vasExam.Col = 4
        vasExam.ColHidden = False
        vasExam.Col = 5
        vasExam.ColHidden = False
'''        vasExam.Col = 6
'''        vasExam.ColHidden = False
        vasExam.Col = 7
        vasExam.ColHidden = False
        vasExam.Col = 11
        vasExam.ColHidden = False
        vasExam.Col = 13
        vasExam.ColHidden = False
    End If

End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long

    lRow = vasID.ActiveRow

    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow + 1
    vasActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
End Sub

Private Sub cmdNotExit_Click()
    spNotTrans.Visible = False
End Sub

Private Sub cmdNotResult_Click()
    ClearSpread vasNotResult

    SQL = "select '', barcode, examcode, examname, result from pat_res " & vbCrLf & _
          "where examdate = '" & Format(dtpNotDate.Value, "yyyymmdd") & "' and sendflag <> 'C'"
    res = db_select_Vas(gLocal, SQL, vasNotResult)

End Sub

Private Sub cmdQC_Click()
'''    frmQCResSch.Show
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer

    Var_Clear

    txtData.Text = ""
    txtErr.Text = ""

    vasID.Row = 1
    vasID.Row2 = vasID.MaxRows
    vasID.Col = 1
    vasID.Col2 = vasID.MaxCols
    vasID.BlockMode = True
    vasID.BackColor = RGB(255, 255, 255)
    vasID.Action = 3
    vasID.BlockMode = False

    ClearSpread vasID
    ClearSpread vasRes
    ClearSpread vasExam
    
    dtpToday = Format(Date, "yyyy/mm/dd")
    
    gRow = 0
End Sub

Private Sub cmdResInsert_Click()
'    Dim iRow As Long
'    Dim iCol As Long
'    Dim lsNewBarcode As String
'    Dim lsOldBarcode As String
'    Dim lsBarcode As String
'
'    Dim rv As Integer
'    Dim i As Long
'
'    iRow = vasExam.DataRowCnt + 1
'
'    If iRow > vasExam.MaxRows Then
'        vasExam.MaxRows = iRow
'    End If
'
'    lsNewBarcode = InputBox("변경할 검체번호를 입력하세요.", "검체번호변경")
'
'    lsBarcode = Left(lsNewBarcode, 11)
'    SQL = "select barcode ,receno ,pid, pname, examcode from pat_res where barcode = '" & lsBarcode & "'"
'    res = db_select_Col(gLocal, SQL)
'
'    If res < 1 Then
'        SQL = "SELECT SPECIMEN_SER, LAB_NO, PATIENT_ID, PATIENT_NAME, EXAM_ITEM_CODE FROM LMI_ORDER " & vbCrLf & _
'              "WHERE SPECIMEN_SER  = '" & lsBarcode & "' AND EXAM_ITEM_CODE IN (" & gAllExam & ")"
'        res = db_select_Col(gServer, SQL)
'        SetText argSpread, Trim(gReadBuf(0)), asRow, colBarcode
'        SetText argSpread, Trim(gReadBuf(1)), asRow, colReceno
'        SetText argSpread, Trim(gReadBuf(2)), asRow, colPID
'        SetText argSpread, Trim(gReadBuf(3)), asRow, colPName
'        SetText argSpread, Trim(gReadBuf(4)), asRow, colExamCode
'
'    Else
'        SetText argSpread, Trim(gReadBuf(0)), asRow, colBarcode
'        SetText argSpread, Trim(gReadBuf(1)), asRow, colReceno
'        SetText argSpread, Trim(gReadBuf(2)), asRow, colPID
'        SetText argSpread, Trim(gReadBuf(3)), asRow, colPName
'        SetText argSpread, Trim(gReadBuf(4)), asRow, colExamCode
'
'    End If
'
'
'
'
'    If res < 1 Then
'    Else
'
'
'
'
'
'        SetText vasExam, lsBarcode, iRow, colBarcode
'        SetText vasExam, gReadBuf(0), iRow, colPID
'        SetText vasExam, gReadBuf(9), iRow, colReceDate
'        SetText vasExam, gReadBuf(32), iRow, colPName
'        SQL = "select examname, equipcode from equipexam where examcode = '" & Trim(gReadBuf(6)) & "'"
'        res = db_select_Col(gLocal, SQL)
'        SetText vasExam, Trim(gReadBuf(0)), iRow, colReceCode
'        SetText vasExam, Trim(gReadBuf(1)), iRow, colEquipCode
'
'        SetText vasExam, Format(Date, "yyyymmdd"), iRow, colExamDate
'        SetText vasExam, Format(Time, "hhmm"), iRow, colExamTime
'        SQL = "Select barcode from pat_res " & vbCrLf & _
'              "WHERE examdate = '" & Trim(GetText(vasExam, iRow, colExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, iRow, colExamTime)) & "' " & vbCrLf & _
'              "  AND equipno = '" & gEquip & "' " & vbCrLf & _
'              "  AND equipcode = '" & Trim(GetText(vasExam, iRow, colEquipCode)) & "'" & vbCrLf & _
'              "  AND barcode = '" & lsBarcode & "' "
'        res = db_select_Col(gLocal, SQL)
'
'        If res > 0 Then
'        Else
'            SQL = "INSERT INTO pat_res (examdate, examtime, equipno, " & _
'                  "barcode, sampletype, receno, " & _
'                  "pid, pname, jumin, page, psex, " & _
'                  "recedate, seqno, diskno, posno, " & _
'                  "equipcode, examcode, " & _
'                  "result, sendflag, examname, " & _
'                  "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, subcode, rececode, bigo, equipres ) " & vbCrLf & _
'                  "VALUES ('" & Trim(GetText(vasExam, iRow, colExamDate)) & "', '" & Trim(GetText(vasExam, iRow, colExamTime)) & "', '" & Trim(gEquip) & "', " & _
'                  "'" & Trim(GetText(vasExam, iRow, colBarcode)) & "','', '', " & _
'                  "'" & Trim(GetText(vasExam, iRow, colPID)) & "', '" & Trim(GetText(vasExam, iRow, colPName)) & "', '', 0, '', " & _
'                  "'" & Trim(GetText(vasExam, iRow, colReceDate)) & "', '', '', '', " & vbCrLf & _
'                  "'" & Trim(GetText(vasExam, iRow, colEquipCode)) & "', '" & Trim(GetText(vasExam, iRow, colExamCode)) & "', " & _
'                  "'', 'B', '" & Trim(GetText(vasExam, iRow, colExamName)) & "', " & vbCrLf & _
'                  "'', '', '', '', " & _
'                  "'', '', '00','" & Trim(GetText(vasExam, iRow, colReceCode)) & "', '" & Trim(GetText(vasExam, iRow, colErrState)) & "', '' ) "
'            res = SendQuery(gLocal, SQL)
'            If res = -1 Then
'                SaveQuery SQL
'                'Exit Function
'            End If
'        End If
'
'    End If

End Sub

Private Sub cmdResSave_Click()
    'Proc_Result txtBarcode
End Sub

Private Sub cmdSend_Click()
    Dim lRow As Long
    Dim sTransTime As String
    
'''    If Trim(txtUser.Text) = "" Then
'''        MsgBox "사용자ID를 입력하세요."
'''        Exit Sub
'''
'''    End If
    
    sTransTime = Format(Time, "hhmm")
    
    For lRow = 1 To vasExam.DataRowCnt
        vasExam.Row = lRow
        vasExam.Col = 1
        If vasExam.Value = 1 Then
            res = Insert_Data(lRow)
'            vasExam.Value = 0
            If res = -1 Then
                SetForeColor vasExam, lRow, lRow, colCheckBox, colCheckBox, 255, 0, 0
                SetText vasExam, "실패", lRow, colState
                Err_Data lRow
            Else
                vasExam.Row = lRow
                vasExam.Col = 1
                
                
                SetBackColor vasExam, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasExam, "완료", lRow, colState
                SetText vasExam, sTransTime, lRow, colTransTime
                
                
                If Trim(GetText(vasExam, lRow, colErrState)) = "Positive" Then
                    spErr.Caption = Trim(GetText(vasExam, lRow, colBarcode)) & " [Positive] 결과"
                    tmErr.Enabled = True
                End If
            
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C', resdate = '" & sTransTime & "' " & vbCrLf & _
                      " Where equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
        End If

        vasExam.Row = lRow
        vasExam.Col = 1
        vasExam.Value = 0
    Next lRow

    TransCheck
End Sub

Private Sub Err_Data(asRow As Long)
    Dim sErrData

    sErrData = ""

    If Trim(GetText(vasExam, asRow, colReceCode)) <> Trim(GetText(vasExam, asRow, colExamCode)) Then
        sErrData = "처방항목과 다른 항목결과입니다."
    End If
    
    If Trim(GetText(vasExam, asRow, colResult)) = "Aborted" Then
        sErrData = "결과값이 [Aborted] 입니다."
    End If
    
    If sErrData <> "" Then
        SQL = "update pat_res set bigo = '" & sErrData & "' " & vbCrLf & _
              "where barcode = '" & Trim(GetText(vasExam, asRow, colBarcode)) & "' " & vbCrLf & _
              "and examdate = '" & Trim(GetText(vasExam, asRow, colExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, asRow, colExamTime)) & "'"
        res = SendQuery(gLocal, SQL)
        SetText vasExam, sErrData, asRow, colErrState
        
        Save_Raw_Data "[SQL" & res & "]" & SQL
    End If
    
    spErr.Caption = Trim(GetText(vasExam, asRow, colBarcode)) & " 결과전송 실패"
    tmErr.Enabled = True

End Sub
Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow - 1
    vasActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1
End Sub

Private Sub cmdUser_Click()
    frmUser.Show 0

End Sub

Private Sub cmdWorkList_Click()
'''    Timer1.Enabled = False
'''    frmWorkList.Show
End Sub

Private Sub Command_close_Click()
    Unload Me
End Sub

Private Sub Command_config_Click()
    frmConfig.Show 1
End Sub

Private Sub Command_setup_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub

Private Sub Command5_Click()
    Dim strSendData As String
    
    XML_Parsing Text2.Text
    H232 gXML.Barcode, gXML.DateTime, gXML.EquipCode, gXML.Result, gXML.StatusCode, gXML.InterpretationCode
    strSendData = WinSock_ACK(gXML.Rece_ControlID, "1")
    Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & "]" & strSendData
    Winsock1.SendData strSendData

'''    WinSock_ACK "1000"
    Text2 = ""
End Sub

'Private Sub Command4_Click()
'    Amplicor_INIT
'End Sub

Private Sub Form_Load()
    Dim sDate As String

    '메인화면 관련
    Me.Left = 0
    Me.Top = 0
    Me.Height = 11190
    Me.Width = 15360

    gResCol = 16

    '변수 초기화
    cmdReset_Click

    'ini파일에서 정보 가져오기
    GetSetup
   
    
'    If Not Connect_Server Then
'        MsgBox "연결되지 않았습니다."
'        cn_Server_Flag = False
'        Exit Sub
'    Else
        cn_Server_Flag = True
'    End If

    '로컬접속
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    
    If cn_Local_Flag = True Then: TransCheck
    
    gTimerReq = 0
    
    WinSock_Listen Winsock1, 1
    'WinSock_Listen Winsock2, 2
    

''    If Text_Today.Text = "2007-07-12" Then
'        SQL = " Alter Table pat_res Alter Column diskno text(20) "
'        res = SendQuery(gLocal, SQL)
'
'        SQL = " Alter Table pat_res Alter Column posno text(20) "
'        res = SendQuery(gLocal, SQL)
''    End If

    '서버접속
'    cn_Server_Flag = dce_setenv("client.env", "", "")

    '검사일자
    dtpToday = Format(CDate(GetDateFull), "yyyy/mm/dd")
    dtpToday_1 = Format(CDate(GetDateFull), "yyyy/mm/dd")

    '====================로컬 DB지우기 - 30일 보관======================
    sDate = Format(DateAdd("y", dtpToday, -30), "yyyymmdd")

    SQL = "Delete from pat_res where examdate < '" & sDate & "' "
    SendQuery gLocal, SQL
    '===================================================================
    
    '검사코드
    GetExamCode
    
    SQL = " Select unit From pat_res "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table pat_res Add Column unit text(20) "
        res = SendQuery(gLocal, SQL)
    End If
    
'    SQL = " Alter Table pat_res Alter Column result text(50) "
'    res = SendQuery(gLocal, SQL)
'
'    SQL = " Alter Table pat_res Alter Column sampletype text(20) "
'    res = SendQuery(gLocal, SQL)
    
    txtBuff = ""
    txtData = ""
    
    vasExam.MaxRows = 0
    spNotTrans.Visible = False

    cmdCol1_Click
    tmErr.Enabled = False
    gSend_ControlID = 10001
    gSend_ControlID2 = 10001
    
'''    State_Check = 0
'''    gWait = 0
    defClr
    gRow = 0
    chkMode.Value = 1
'    Amplicor_INIT
    
'    Timer1.Enabled = False
    
End Sub

Sub PatInfo(argSpread As vaSpread, asSpecID As String, asRow As Long)
    
    Dim lsBarcode As String
    
    Dim rv As Integer
    Dim i As Long
    
'    ClearSpread vasCode
    
    lsBarcode = asSpecID
'    SQL = "p_interfacequery '1', '" & lsBarcode & "'"
'    res = db_select_Col(gServer, SQL)
    SQL = "select barcode ,receno ,pid, pname, examcode from pat_res where barcode = '" & lsBarcode & "'"
    res = db_select_Col(gLocal, SQL)
    
                                                               
    If res < 1 Then
'''        SQL = "SELECT SPECIMEN_SER, LAB_NO, PATIENT_ID, PATIENT_NAME, EXAM_ITEM_CODE FROM LMI_ORDER " & vbCrLf & _
'''              "WHERE SPECIMEN_SER  = '" & lsBarcode & "' AND EXAM_ITEM_CODE IN (" & gAllExam & ")"
'''        select a.bunho, b.suname
'''  from itf1001 a, medi.VW_ITF_OUT0101 b
''' Where a.bunho = b.bunho
'''   and a.fkocs = 바코드값
        SQL = "SELECT A.FKOCS, A.BUNHO, A.BUNHO, B.SUNAME, A.HANGMOG_CODE FROM ITF.ITF1001 A, MEDI.VW_ITF_OUT0101 B " & vbCrLf & _
              "WHERE A.BUNHO = B.BUNHO AND A.FKOCS = '" & lsBarcode & "' AND A.HANGMOG_CODE IN (" & gAllExam & ")"
        res = db_select_Col(gServer, SQL)
'''        SetText argSpread, Trim(gReadBuf(0)), asRow, colBarcode
        'SetText argSpread, Trim(gReadBuf(1)), asRow, colReceno
        SetText argSpread, Trim(gReadBuf(2)), asRow, colPID
        SetText argSpread, Trim(gReadBuf(3)), asRow, colPName
        SetText argSpread, Trim(gReadBuf(4)), asRow, colExamCode
        SetText argSpread, Trim(gReadBuf(4)), asRow, colReceCode
    Else
'''        SetText argSpread, Trim(gReadBuf(0)), asRow, colBarcode
        'SetText argSpread, Trim(gReadBuf(1)), asRow, colReceno
        SetText argSpread, Trim(gReadBuf(2)), asRow, colPID
        SetText argSpread, Trim(gReadBuf(3)), asRow, colPName
        SetText argSpread, Trim(gReadBuf(4)), asRow, colExamCode
        SetText argSpread, Trim(gReadBuf(4)), asRow, colReceCode
    End If
    
End Sub

Function EquipExamCode(argEquipCode As String, argPID As String, argSENO As String, argSEQN As String) As String
'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
    Dim i As Integer
    Dim sExamCode As String

    EquipExamCode = ""

    If Trim(argEquipCode) = "" Then
        Exit Function
    End If

    ClearSpread vasTemp1
    sExamCode = ""

    SQL = " Select examcode From EquipExam " & vbCrLf & _
          " Where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          " And equipcode = '" & Trim(argEquipCode) & "' "
    res = db_select_Vas(gLocal, SQL, vasTemp1)

    If vasTemp1.DataRowCnt < 1 Then
        Exit Function
    End If

    For i = 1 To vasTemp1.DataRowCnt
        If sExamCode <> "" Then
            sExamCode = sExamCode & ",'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        Else
            sExamCode = "'" & Trim(GetText(vasTemp1, i, 1)) & "'"
        End If
    Next i

    SQL = " Select SUCD From LRESULT " & CR & _
          " Where PAID = '" & Trim(argPID) & "' " & vbCrLf & _
          "   and SENO = " & argSENO & vbCrLf & _
          "   and SEQN = " & argSEQN & vbCrLf & _
          "   and SUCD in ( " & sExamCode & ")  "

    res = db_select_Col(gServer, SQL)

    If gReadBuf(0) <> "" Then
        EquipExamCode = Trim(gReadBuf(0))
    End If
    
End Function

Function GetExamCode() As Integer
    Dim i As Long
    Dim j As Long
    
    gAllExam = ""
    
    ClearSpread vasTemp
    GetExamCode = -1
    
    '장비코드,검사코드,검사명,참고치_Low,참고치_High,오더구분
    SQL = "Select equipcode, examcode, examname, reflow, refhigh, ordgubun, subcode " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  seqno "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    If res > 0 Then
        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 8)
    Else
        SaveQuery SQL
        Exit Function
    End If
         
'''    vasExam.MaxCols = vasTemp.DataRowCnt * 5 + colResult - 1
''''    SetText vasExam, "COPY", 0, vasExam.MaxCols
'''    colResult1 = vasTemp.DataRowCnt * 4 + colResult
    
''''    vasExam.ColWidth(13) = 0
'''    vasExam.ColWidth(14) = 0
'''    vasExam.ColWidth(15) = 0
'''    vasExam.ColWidth(16) = 0
'''    vasExam.ColWidth(17) = 0
    
    For i = 1 To vasTemp.DataRowCnt
        gArrEquip(i, 1) = i
        For j = 1 To 7
            gArrEquip(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
        
        '검사명 디스플레이
        'MsgBox colResult + (i - 1) * 4
'''        SetText vasExam, gArrEquip(i, 4), 0, colResult + (i - 1) * 4
'''        vasExam.ColWidth(colResult + (i - 1) * 4) = 15
'''        vasExam.ColWidth(colResult + (i - 1) * 4 + 1) = 0
'''        vasExam.ColWidth(colResult + (i - 1) * 4 + 2) = 0
'''        vasExam.ColWidth(colResult + (i - 1) * 4 + 3) = 0
'''
'''        vasExam.ColWidth(colResult1 + i - 1) = 0
        
        If gAllExam = "" Then
            gAllExam = "'" & Trim(GetText(vasTemp, i, 2)) & "'"
        Else
            gAllExam = gAllExam & ",'" & Trim(GetText(vasTemp, i, 2)) & "'"
        End If
        
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Local
    DisConnect_Server
    
    Unload Me
    
    End
    
End Sub

Public Function HConnectState(asData As String) As Boolean
    Dim lsData As String
    Dim i As Integer
    
    HConnectState = False
    
    lsData = asData
    i = InStr(1, lsData, chrSTX)
    
    If i > 0 Then
        lsData = Mid(lsData, i + 1)
        i = InStr(1, lsData, chrHT)
        If i > 0 Then
            lsData = Mid(lsData, i + 1)
            i = InStr(1, lsData, chrHT)
            If i > 0 Then
                lsData = Mid(lsData, 1, i - 1)
                If lsData = "00FD" Or lsData = "00FF" Then
                    HConnectState = True
                End If
            End If
        End If
    End If
End Function

Public Function HSerial(asData As String) As String

    Dim lsData As String
    Dim i As Integer
    
    HSerial = ""
    
    lsData = asData
    i = InStr(1, lsData, chrSTX)
    
    If i > 0 Then
        lsData = Mid(lsData, i + 1)
        i = InStr(1, lsData, chrHT)
        If i > 0 Then
            lsData = Mid(lsData, i + 1)
            i = InStr(1, lsData, chrHT)
            If i > 0 Then
                lsData = Mid(lsData, 1, i - 1)
                HSerial = lsData
            End If
        End If
    End If
End Function

Private Function Make_Msg(asMsgState As Integer, asMsgSignal As String) As String

End Function

Private Sub H232(asBarcode As String, asDateTime As String, asEquipCode As String, asResult As String, _
                 asStateCode As String, asRefCode As String)

    Dim sData As String
    Dim i As Integer
    Dim sStr(1 To 30) As String
    Dim j As Long
    Dim k As Long
    Dim iRow As Integer
    
    Dim sBarcode As String
    Dim sEquipCode As String
    Dim sResult As String
    Dim sEquipRes As String
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sSeqNo As String
    Dim sResType As String
    Dim sResPoint As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim liRet
    Dim sGiho As String
    Dim sFlag As String
    Dim sResFlag As String
    
    On Error GoTo ErrRes:

    sBarcode = asBarcode
       
    sEquipCode = asEquipCode
    sResult = asResult
    sEquipRes = sResult
    
    '2009-06-02T18:22:18+00:00
    
    sExamDate = Format(Mid(asDateTime, 1, 10), "yyyymmdd")
    
    sExamTime = Format(Mid(asDateTime, 12, 5), "hhmm")
    sGiho = asRefCode
    sFlag = asStateCode
    
    gSpecID = sBarcode
    
    SQL = "select barcode from pat_res where barcode = '" & sBarcode & "'  and examdate = '" & sExamDate & "' and examtime = '" & sExamTime & "'"
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) = sBarcode Then
        Exit Sub
    End If
    
    gReadBuf(0) = ""
    
    SQL = "select barcode from worklist where barcode = '" & sBarcode & "' and resdatetime = '" & sExamTime & "'"
    res = db_select_Col(gLocal, SQL)

    If Trim(gReadBuf(0)) = sBarcode Then
        Exit Sub
    End If
    
    SQL = "select examcode from equipexam where equipcode = '" & sEquipCode & "' "
    res = db_select_Col(gLocal, SQL)

    If Trim(gReadBuf(0)) = "" Then
        Exit Sub
    End If
    
    
    glRow = -1
    For i = 1 To vasExam.DataRowCnt
        If Trim(GetText(vasExam, i, colBarcode)) = gSpecID And Trim(GetText(vasExam, i, colExamDate)) = sExamDate And Trim(GetText(vasExam, i, colExamTime)) = sExamTime Then
            glRow = i
            
            Exit For
        End If
    Next i

    If glRow = -1 Then  ' vasID에 없는 검체의 결과가 나올 때 데이터 추가
        glRow = vasExam.DataRowCnt + 1
        If glRow > vasExam.MaxRows Then
            vasExam.MaxRows = glRow + 1
        End If
    End If
    
    vasExam.RowHeight(glRow) = 25
    
    SetText vasExam, gSpecID, glRow, colBarcode
    SetText vasExam, sExamDate, glRow, colExamDate
    SetText vasExam, sExamTime, glRow, colExamTime
    

    If Trim(GetText(vasExam, glRow, colPID)) = "" Then
        PatInfo vasExam, gSpecID, glRow
    End If
    gPreSpecID = gSpecID
    gPreRow = glRow
        
    SetText vasExam, "결과", glRow, colState
    SetBackColor vasExam, glRow, glRow, colCheckBox, colState, 255, 250, 205

    sExamCode = ""
    sExamName = ""
    
    gReadBuf(0) = ""
    SQL = "Select EquipCode, ExamCode, ExamName, SeqNo, RSGubun, resprec, RefLow, RefHigh " & vbCrLf & _
          "from equipexam where equipno = '" & gEquip & "' and EquipCode = '" & sEquipCode & "' "
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) = sEquipCode Then
        sExamCode = Trim(gReadBuf(1))
        sExamName = Trim(gReadBuf(2))
        sSeqNo = Trim(gReadBuf(3))
        sResType = Trim(gReadBuf(4))
        sResPoint = Trim(gReadBuf(5))
        sRefLow = Trim(gReadBuf(6))
        sRefHigh = Trim(gReadBuf(7))
    Else
        sEquipCode = ""
    End If
    
    
    If sResPoint = "0" Then
        sResult = Format(sResult, "#0")
    ElseIf sResPoint = "1" Then
        sResult = Format(sResult, "#0.0")
    ElseIf sResPoint = "2" Then
        sResult = Format(sResult, "#0.00")
    ElseIf sResPoint = "3" Then
        sResult = Format(sResult, "#0.000")
    ElseIf sResPoint = "4" Then
        sResult = Format(sResult, "#0.0000")
    Else
        sResult = sResult
    End If
    
    If IsNumeric(sResult) = True And IsNumeric(sRefLow) = True Then
        If CCur(sResult) < CCur(sRefLow) Then
              sResult = sRefLow
        End If
        
    End If
    
    If IsNumeric(sResult) = True And IsNumeric(sRefHigh) = True Then
        If CCur(sResult) > CCur(sRefHigh) Then
            sResult = sRefHigh
        End If
        
    End If
    
    If sGiho = "<" Then
        sResult = "<" & sResult
    ElseIf sGiho = ">" Then
        sResult = ">" & sResult
    End If
    
    If sFlag <> "A" Then
        sResult = "Aborted"
    End If
    
    If InStr(1, sEquipRes, "0.030 - 0.100") > 0 Then
        sResult = "0.03"
    End If
    
    
    SetText vasExam, sResult, glRow, colResult
    SetText vasExam, sEquipRes, glRow, colEquipRes
    SetText vasExam, sEquipCode, glRow, colEquipCode
    SetText vasExam, sExamCode, glRow, colExamCode
    SetText vasExam, sExamName, glRow, colExamName
    
    If sResFlag = "Positive" Then
        SetText vasExam, sResFlag, glRow, colErrState
        SetForeColor vasExam, glRow, glRow, colErrState, colErrState, 255, 0, 0
        
    End If
    
    SQL = "Select barcode from pat_res " & vbCrLf & _
          "WHERE examdate = '" & sExamDate & "' and examtime = '" & sExamTime & "' " & vbCrLf & _
          "  AND equipno = '" & gEquip & "' " & vbCrLf & _
          "  AND equipcode = '" & sEquipCode & "'" & vbCrLf & _
          "  AND barcode = '" & sBarcode & "' "
    res = db_select_Col(gLocal, SQL)
    
    If res > 0 Then
    Else
        SQL = "INSERT INTO pat_res (examdate, examtime, equipno, " & _
              "barcode, sampletype, receno, " & _
              "pid, pname, jumin, page, psex, " & _
              "recedate, seqno, diskno, posno, " & _
              "equipcode, examcode, " & _
              "result, sendflag, examname, " & _
              "refflag,panicflag, deltaflag, unit, refvalue, panicvalue, subcode, rececode, bigo, equipres ) " & vbCrLf & _
              "VALUES ('" & sExamDate & "', '" & sExamTime & "', '" & Trim(gEquip) & "', " & _
              "'" & Trim(GetText(vasExam, glRow, colBarcode)) & "','', '', " & _
              "'" & Trim(GetText(vasExam, glRow, colPID)) & "', '" & Trim(GetText(vasExam, glRow, colPName)) & "', '', 0, '', " & _
              "'" & Trim(GetText(vasExam, glRow, colReceDate)) & "', '" & sSeqNo & "', '', '', " & vbCrLf & _
              "'" & sEquipCode & "', '" & sExamCode & "', " & _
              "'" & sResult & "', 'B', '" & sExamName & "', " & vbCrLf & _
              "'', '', '', '', " & _
              "'', '', '00','" & Trim(GetText(vasExam, glRow, colReceCode)) & "', '" & Trim(GetText(vasExam, glRow, colErrState)) & "', '" & sEquipRes & "' ) "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            'Exit Function
        End If
    End If

    If chkMode.Value = 1 Then
        Dim sTTime As String
        
        sTTime = Format(Time, "hhmm")
        
        liRet = -1
        If glRow < 1 Then
            Exit Sub
        End If

        If Trim(GetText(vasExam, glRow, colPID)) <> "" Then
            liRet = Insert_Data(glRow)
        End If

        If liRet = -1 Then
            SetBackColor vasExam, glRow, glRow, colState, colState, 255, 0, 0
            SetText vasExam, "실패", glRow, colState
            Err_Data glRow
        Else
            SetBackColor vasExam, glRow, glRow, 1, colState, 202, 255, 112
            SetText vasExam, "완료", glRow, colState
            SetText vasExam, sTTime, glRow, colTransTime
            

            If Trim(GetText(vasExam, glRow, colErrState)) = "Positive" Then
                spErr.Caption = Trim(GetText(vasExam, glRow, colBarcode)) & " [Positive] 결과"
                tmErr.Enabled = True
            End If


            'Local 상태를 서버전송(C)로 바꿈
            SQL = " Update pat_res Set " & vbCrLf & _
                  " sendflag = 'C', resdate = '" & sTTime & "' " & vbCrLf & _
                  " Where equipno = '" & gEquip & "' " & vbCrLf & _
                  " And barcode = '" & Trim(GetText(vasExam, glRow, colBarcode)) & "' " '& vbCrLf & _
                  " And equipcode = '" & sEquipCode & "' "
            res = SendQuery(gLocal, SQL)
            If res = -1 Then
                SaveQuery SQL
                Exit Sub
            End If
        End If

    End If
     TransCheck

ErrRes:
     Exit Sub
End Sub

Sub Save_Raw_Data(argSQL As String)
'argSQL의 내용을 파일로 저장
    Dim FilNum
    Dim sFileName As String
    
    FilNum = FreeFile
    
    If Dir(App.Path & "\Log", vbDirectory) <> "Log" Then
        MkDir (App.Path & "\Log")
    End If
    
    sFileName = Format(dtpToday, "yyyymmdd")
    
    Open App.Path & "\Log\" & sFileName & ".txt" For Append As FilNum
    Print #FilNum, argSQL
    Close FilNum
End Sub

Function Insert_Data(ByVal argSpcRow As Integer) As Integer
'서버의 데이타 베이스에 저장
    Dim iRow As Integer
    Dim i As Integer
    Dim j As Integer

'    Dim clsResult As New clsResult
    Dim sBarcode, sTstcode As String
    Dim sPID As String
    Dim rc As Integer
    Dim mCnt As Integer
    Dim oerrmsg$
    Dim ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$()
    Dim sSampleSeq As String
    Dim sSampleDate As String
    Dim sChartNo As String
    Dim sPart As String
    Dim sSubCode As String
    Dim sExamCode As String
    Dim sResult As String
    Dim sSendFlag As String
    Dim sExamFlag As Boolean
    Dim sTransTime As String
    Dim sExamTime As String
    
    Dim sTransYN As Boolean
    Dim sPName As String
    Dim ss As Integer
    Dim sProcRes As String
    Dim sInsertTime As String
    Dim sTransSeq As String
    Insert_Data = -1
      
    sBarcode = Trim(GetText(vasExam, argSpcRow, colBarcode))
    sExamTime = Trim(GetText(vasExam, argSpcRow, colExamTime))
    sTransTime = Format(Time, "hhmm")
    
'    sPID = Trim(GetText(vasID, argSpcRow, colPID))
    
    'Local에서 환자별로 결과값 가져오기
    ClearSpread vasResTemp
    
    SQL = " Select equipcode, examcode, result, receno, pid, sampletype, recedate, subcode " & vbCrLf & _
          " From pat_res " & vbCrLf & _
          " Where equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & sBarcode & "' and examdate = '" & Trim(GetText(vasExam, argSpcRow, colExamDate)) & "' " & _
          " and examtime = '" & Trim(GetText(vasExam, argSpcRow, colExamTime)) & "' " & _
          " and examcode = '" & Trim(GetText(vasExam, argSpcRow, colExamCode)) & "' "
          
    res = db_select_Vas(gLocal, SQL, vasResTemp)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If

    For i = 1 To vasResTemp.DataRowCnt
        If IsNumeric(Trim(GetText(vasResTemp, i, 3))) = False Then
            If Trim(GetText(vasResTemp, i, 3)) = "Aborted" Then
                Exit Function
            End If
        End If
    Next

    'ClearSpread vasServerTemp
        
'''    Connect_Server

    cn_Ser.BeginTrans
    For i = 1 To vasResTemp.DataRowCnt
        sTransTime = Format(Date, "yyyymmdd") & Format(Time, "hhmm")
        sInsertTime = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
        sTstcode = Trim(GetText(vasResTemp, i, 2))
        sResult = Trim(GetText(vasResTemp, i, 3))
        
        '///// 오더테이블에 접수된 검사가 있는지 확인
'        SQL = "SELECT EXAM_ITEM_CODE FROM LMI_ORDER " & _
'              " WHERE SPECIMEN_SER = '" & sBarcode & "' " & _
'              "   AND EXAM_ITEM_CODE = '" & Trim(GetText(vasResTemp, i, 2)) & "' "
'        res = db_select_Col(gServer, SQL)
'        If Trim(gReadBuf(0)) = "" Then: gReadBuf(0) = "": cn_Ser.RollbackTrans: Exit Function
        '//////////
        
        SQL = "UPDATE ITF.ITF1001 " & vbCrLf & _
              "SET RSTFLAG = '3', RSTDT = TRUNC(SYSDATE), VIEWRST = '" & sResult & "' " & vbCrLf & _
              "WHERE FKOCS = '" & sBarcode & "' AND HANGMOG_CODE = '" & sTstcode & "' AND SPCFLAG = '2'"
        res = SendQuery(gServer, SQL)
        
        
        If res < 0 Then
            cn_Ser.RollbackTrans
            SaveQuery SQL
            'DisConnect_Server
            Exit Function
            
        End If
    Next
    Insert_Data = 1
    cn_Ser.CommitTrans
'''    DisConnect_Server
    
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

Private Sub spErr_Click()
    spErr.Caption = ""
'''    SQL = "select barcode from worklist where barcode = '" & sBarcode & "' and resdatetime = '" & sExamTime & "'"
    
''    res = db_select_Col(gLocal, SQL)
''    If Trim(gReadBuf(0)) = sBarcode Then
''    Else
''        SQL = "insert into worklist(barcode, ResDateTime) values('" & sBarcode & "', '" & sExamTime & "')"
''        res = SendQuery(gLocal, SQL)
''    End If
    
            
    tmErr.Enabled = False
End Sub

'''Private Sub spMissResNow_DblClick()
'''    dtpNotDate.Value = Format(Date, "yyyy-mm-dd")
'''    ClearSpread vasNotResult
'''    spNotTrans.Visible = True
'''End Sub
'''
'''Private Sub spMissResPast_DblClick()
'''    dtpNotDate.Value = Format(Date - 1, "yyyy-mm-dd")
'''    ClearSpread vasNotResult
'''    spNotTrans.Visible = True
'''End Sub

Private Sub Text_Today_GotFocus()
    SelectFocus Text_Today
End Sub

Private Sub Text_Today_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCall_Click
    ElseIf KeyCode = vbKeyF7 Then
'''        frmQCResSch.Show
    End If
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tm_H232_2_Timer()

'''    If MSComm2.PortOpen = False And gSetup2.gPort <> 0 Then
'''        MSComm2.PortOpen = True
'''    Else
'''        Exit Sub
'''
'''    End If
'''
'''    If H232_Connect_state_2 = False Then
'''        MSComm2.Output = H232_Function_2(H232_Connect)
'''    End If
End Sub

Private Sub tm_H232_Timer()

'''    If MSComm1.PortOpen = False And gSetup.gPort <> 0 Then
'''        MSComm1.PortOpen = True
'''    Else
'''        Exit Sub
'''
'''    End If
'''
'''    If H232_Connect_state = False Then
'''        MSComm1.Output = H232_Function(H232_Connect)
'''    End If
'''
End Sub

Private Sub tmErr_Timer()
    Beep
    
End Sub

Private Sub defClr()
    SSPanel6.BackColor = &H800000
    spMissResNow.BackColor = &H800000
    spMissResNow.Caption = 0
    spMissResPast.Caption = 0
    SSPanel4.BackColor = &H800000
    spMissResPast.BackColor = &H800000
End Sub

Private Sub TransCheck()
    SQL = "select count(receno) from pat_res where examdate = '" & Format(Date, "yyyymmdd") & "' and sendflag <> 'C' and pname <> '' "
    res = db_select_Col(gLocal, SQL)


    If IsNumeric(Trim(gReadBuf(0))) = True Then
        spMissResNow.Caption = Trim(gReadBuf(0))
        If Trim(gReadBuf(0)) = "0" Then
            SSPanel6.BackColor = &H800000
            spMissResNow.BackColor = &H800000
        Else
            SSPanel6.BackColor = &HFF&
            spMissResNow.BackColor = &HFF&
        End If

    Else
        SSPanel6.BackColor = &H800000
        spMissResNow.BackColor = &H800000
        spMissResNow.Caption = 0
    End If

    SQL = "select count(*) from pat_res "
    SQL = SQL & vbCrLf & "where examdate between '" & Format(dtpToday_1, "yyyymmdd") & "' AND '" & Format(dtpToday - 1, "yyyymmdd") & "' "
    SQL = SQL & vbCrLf & "  and sendflag <> 'C' and pname <> '' "
    res = db_select_Col(gLocal, SQL)

    If IsNumeric(Trim(gReadBuf(0))) = True Then
        spMissResPast.Caption = Trim(gReadBuf(0))
        If Trim(gReadBuf(0)) = "0" Then
            SSPanel4.BackColor = &H800000
            spMissResPast.BackColor = &H800000
        Else
            SSPanel4.BackColor = &HFF&
            spMissResPast.BackColor = &HFF&
        End If
    Else

        spMissResPast.Caption = 0
        SSPanel4.BackColor = &H800000
        spMissResPast.BackColor = &H800000
    End If

End Sub

Private Sub tmResRequest_Timer()

End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sID As String

    If KeyCode = 8 Then
        txtUser.Text = ""
        lblUser.Caption = ""
        lblUser.Caption = ""
        Exit Sub
        
    ElseIf KeyCode = 13 Then
'        sID = Trim(txtUser.Text) & Chr(KeyCode)
'        txtUser.Text = sID
        lblUser.Caption = ""
        SQL = "select user_name from pword where user_id = '" & Trim(txtUser.Text) & "'"
        res = db_select_Col(gServer, SQL)
        If res > 0 Then
            lblUser.Caption = Trim(gReadBuf(0))
        Else
            MsgBox "잘못된 사용자 ID 입니다."
            txtUser.Text = ""
            Exit Sub
        End If
    End If
    
    
    
End Sub

Private Sub vasExam_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim iRow As Long
    Dim iCol As Long
    Dim lsNewBarcode As String
    Dim lsOldBarcode As String
    Dim lsBarcode As String
    
    Dim rv As Integer
    Dim i As Long
    
    iRow = Row
    iCol = Col
    
    If iCol = colBarcode Then
        
        UserState = False
        frmMod.Show 1
        If UserState = False Then
            Exit Sub
        End If
        

        lsOldBarcode = Trim(GetText(vasExam, iRow, colBarcode))
        lsNewBarcode = InputBox("변경할 검체번호를 입력하세요.", "검체번호변경")
        
'''        SQL = "select barcode from pat_res where barcode = '" & lsNewBarcode & "'"
'''        res = db_select_Col(gLocal, SQL)
'''
'''        If Trim(gReadBuf(0)) = lsNewBarcode Then
'''            MsgBox "이미 입력된 바코드 번호입니다. "
'''            Exit Sub
'''        End If
        
        If Trim(lsNewBarcode) <> "" Then
            lsBarcode = Left(lsNewBarcode, 11)
            SQL = "p_interfacequery '1', '" & lsBarcode & "'"
            res = db_select_Col(gServer, SQL)
            
            If res < 1 Then
                SQL = "update pat_res set barcode = '" & lsNewBarcode & "' " & vbCrLf & _
                      "where equipno = '" & gEquip & "' and barcode = '" & lsOldBarcode & "' " & vbCrLf & _
                      "and examdate = '" & Trim(GetText(vasExam, iRow, colExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, iRow, colExamTime)) & "'"
                res = SendQuery(gLocal, SQL)
            
            Else
                SQL = "update pat_res set barcode = '" & lsNewBarcode & "', pid = '" & Trim(gReadBuf(0)) & "', " & vbCrLf & _
                      "recedate = '" & Trim(gReadBuf(9)) & "', receno = '', " & vbCrLf & _
                      "seqno = '" & Trim(gReadBuf(11)) & "', pname = '" & Trim(gReadBuf(32)) & "' " & vbCrLf & _
                      "where equipno = '" & gEquip & "' and barcode = '" & lsOldBarcode & "' " & vbCrLf & _
                      "and examdate = '" & Trim(GetText(vasExam, iRow, colExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, iRow, colExamTime)) & "'"
                res = SendQuery(gLocal, SQL)
                
                Save_Raw_Data "[SQL]" & SQL
                
                SetText vasExam, gReadBuf(0), iRow, colPID
                SetText vasExam, gReadBuf(9), iRow, colReceDate
'''                SetText vasExam, gReadBuf(10), iRow, colReceno
'''                SetText vasExam, gReadBuf(11), iRow, colSeqNo
                SetText vasExam, gReadBuf(32), iRow, colPName
                SetText vasExam, lsNewBarcode, iRow, colBarcode
                
            End If
            
        End If
    End If
    
End Sub

Private Sub vasExam_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim lRow As Long
    Dim lCol As Long
    Dim lsID As String
    Dim sAnalyzerID As String
    Dim k As Integer
    Dim sEquipCode As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sSeqNo As String
    Dim sResType As String
    Dim sResPoint As String
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sResult As String
    Dim sErrData As String
    Dim sExamDate As String
    Dim sExamTime As String
    Dim sBarcode As String

    lRow = vasExam.ActiveRow
    lCol = vasExam.ActiveCol
    
    If KeyCode = vbKeyReturn Then
    
        UserState = False
        frmMod.Show 1
        If UserState = False Then
            Exit Sub
        End If
        
        
        gReadBuf(0) = ""
        sResult = Trim(GetText(vasExam, lRow, colResult))
        sErrData = Trim(GetText(vasExam, lRow, colErrState))
        sExamDate = Trim(GetText(vasExam, lRow, colExamDate))
        sExamTime = Trim(GetText(vasExam, lRow, colExamTime))
        sBarcode = Trim(GetText(vasExam, lRow, colBarcode))
               
        If Trim(sResult) = "" Then
            SQL = "select barcode from worklist where barcode = '" & sBarcode & "' and resdatetime = '" & sExamTime & "'"
            res = db_select_Col(gLocal, SQL)
            If Trim(gReadBuf(0)) = sBarcode Then
            Else
                SQL = "insert into worklist(barcode, ResDateTime) values('" & sBarcode & "', '" & sExamTime & "')"
                res = SendQuery(gLocal, SQL)
            End If
            
            SQL = "delete from pat_res " & vbCrLf & _
                  "where equipno = '" & gEquip & "' and barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' " & vbCrLf & _
                  "and examdate = '" & Trim(GetText(vasExam, lRow, colExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, lRow, colExamTime)) & "' " & vbCrLf & _
                  "and examcode = '" & Trim(GetText(vasExam, lRow, colExamCode)) & "'"
            res = SendQuery(gLocal, SQL)
            DeleteRow vasExam, lRow, lRow
            
        Else
            SQL = "update pat_res set result = '" & sResult & "', bigo = '" & sErrData & "' " & vbCrLf & _
                  "where equipno = '" & gEquip & "' and barcode = '" & Trim(GetText(vasExam, lRow, colBarcode)) & "' " & vbCrLf & _
                  "and examdate = '" & Trim(GetText(vasExam, lRow, colExamDate)) & "' and examtime = '" & Trim(GetText(vasExam, lRow, colExamTime)) & "' " & vbCrLf & _
                  "and examcode = '" & Trim(GetText(vasExam, lRow, colExamCode)) & "'"
            res = SendQuery(gLocal, SQL)
        End If
        
        
    
    End If
End Sub

Private Sub vasExam_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row < 1 Or Row > vasExam.DataRowCnt Then
        Exit Sub
    End If
    
    'PopupMenu mnuPop
End Sub

'Private Sub subUp_Click()
'    Dim sValue As String
'    Dim sTmp As String
'    Dim i As Integer
'    Dim j As Integer
'
'    sTmp = ""
'
'    vasID.Row = vasID.ActiveRow
'    vasID.Col = colBarcode
'
'    sTmp = vasID.Text
'
'    sValue = InputBox("변경할 검체번호를 입력하세요")
'
'    If Trim(sValue) <> "" Then
'        If MsgBox("" & sTmp & "를 " & sValue & "로 수정하시겠습니까?", vbYesNo, "확인") = vbYes Then
'
'
'        SQL = "SELECT SPECIMEN_SER, LAB_NO, PATIENT_ID, PATIENT_NAME, EXAM_ITEM_CODE FROM LMI_ORDER " & vbCrLf & _
'              "WHERE SPECIMEN_SER  = '" & sValue & "' AND EXAM_ITEM_CODE IN (" & gAllExam & ")"
'        res = db_select_Col(gServer, SQL)
'
'        If res > 0 Then
'            SetText argSpread, Trim(gReadBuf(0)), asRow, colBarcode
'            SetText argSpread, Trim(gReadBuf(2)), asRow, colPID
'            SetText argSpread, Trim(gReadBuf(3)), asRow, colPName
'
'        Else
'            SetText argSpread, "", asRow, colPID
'            SetText argSpread, "", asRow, colPName
'
'        End If
'
'            '//// 수정해야됨
'            SetText vasID, sValue, vasID.Row, vasID.Col
'
'            If Trim(GetText(vasID, vasID.Row, colBarcode)) <> "" Then
'                PatInfo vasID.Row
'                Call DELETE_LOCAL_ONE(Trim(sTmp), Format(GetText(vasID, vasID.Row, colExamDate), "yyyymmdd"))
'            End If
'        End If
'    End If
'End Sub

Function DELETE_LOCAL_ONE(asBarcode As String, asExamdate As String)
    SQL = "DELETE FROM PAT_RES "
    SQL = SQL & vbCrLf & " WHERE EXAMDATE = '" & asExamdate & "' "
    SQL = SQL & vbCrLf & "   AND EQUIPNO = '" & gEquip & "' "
    SQL = SQL & vbCrLf & "   AND BARCODE = '" & Trim(asBarcode) & "' "
    res = SendQuery(gLocal, SQL)
    
    If res = -1 Then
        SaveQuery SQL
        Exit Function
    End If
End Function

Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        vasSort vasID, Col
    End If
    
    If Row < 0 Or Row > vasID.DataRowCnt Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    
    If Row = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    ElseIf Row = vasID.DataRowCnt Then
        cmdUp.Enabled = True
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
End Sub

Private Sub vasID_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim lsCnt As String
    Dim lsID As String
    Dim lsDate As String
    
    Dim iRow As Integer
    
    'cmdCall_Click
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    lsID = Trim(GetText(vasID, Row, colBarcode))
    
    'Local에서 불러오기
    ClearSpread vasRes

    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "Select a.equipcode, a.examcode, b.examname, a.result, a.result1, b.seqno " & vbCrLf & _
          "from pat_res a, equipexam b " & vbCrLf & _
          "where a.examdate = '" & Format(dtpToday, "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and b.equipno = a.equipno " & vbCrLf & _
          "  and b.equipcode = a.equipcode " & vbCrLf & _
          "  and b.examcode = a.examcode " & vbCrLf & _
          "order by b.seqno, a.equipcode "
    res = db_select_Vas(gLocal, SQL, vasRes)
    'SaveQuery SQL
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasRes.MaxRows = vasRes.DataRowCnt
    'vasSort vasRes, 5, 2
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Integer
    Dim lsID As String
    
    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
            
        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        lsID = Trim(GetText(vasID, iRow, colBarcode))
            
        SQL = " Delete From pat_res " & vbCrLf & _
              " Where examdate = '" & Format(dtpToday, "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & lsID & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasID, iRow, iRow
        ClearSpread vasRes
    End If
End Sub

Function Save_Local_QC(asExamdate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
    Dim sResDateTime As String
    Dim sControl As String
    Dim sLotNo As String
    
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sRefFlag As String
    
    Dim sCnt As String
    
    sResDateTime = Format(CDate(asExamdate), "yyyymmdd hhnnss")
    'sControl = Trim(Left(asBarcode, 2))
    'sLotNo = Trim(Mid(asBarcode, 3))
    sControl = asBarcode
    sRefFlag = ""
    
    SQL = "Select t_mean, t_sd from qcexam " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and validstart >= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and valiend <= '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Col(gLocal, SQL)
    If res > 0 Then
        If IsNumeric(gReadBuf(0)) And IsNumeric(gReadBuf(1)) Then
            sRefLow = CCur(gReadBuf(0)) - CCur(gReadBuf(1))
            sRefHigh = CCur(gReadBuf(0)) + CCur(gReadBuf(1))
            If CCur(sRefHigh) < CCur(asRes2) Then
                sRefFlag = "H"
            End If
            If CCur(sRefLow) > CCur(asRes2) Then
                sRefFlag = "L"
            End If
        End If
    End If
    
    sCnt = ""
    SQL = "Select count(*) from qc_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
          "  and examtime = '" & Mid(sResDateTime, 10, 6) & "' " & vbCrLf & _
          "  and levelname = '" & sControl & "' " & vbCrLf & _
          "  and equipcode = '" & asExamCode & "' "
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        db_RollBack gLocal
        Exit Function
    End If
    res = db_select_Var(gLocal, SQL, sCnt)
    If res <= 0 Then
        SaveQuery SQL
        Exit Function
    End If
    If Not IsNumeric(sCnt) Then sCnt = "0"
    
    If CInt(sCnt) > 0 Then
        SQL = "delete from qc_res " & vbCrLf & _
              "where equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examdate = '" & Left(sResDateTime, 8) & "' " & vbCrLf & _
              "  and examtime = '" & Mid(sResDateTime, 9, 4) & "' " & vbCrLf & _
              "  and levelname = '" & sControl & "' " & vbCrLf & _
              "  and equipcode = '" & asExamCode & "' "
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            'db_RollBack gLocal
            SaveQuery SQL
            Exit Function
        End If
    End If
    SQL = "Insert into qc_res (equipno, examdate, examtime, levelname, equipcode, sresult, result, resflag, remark, examuid, lotno) " & vbCrLf & _
          "values ('" & gEquip & "', '" & Left(sResDateTime, 8) & "', '" & Mid(sResDateTime, 10, 4) & "', '" & sControl & "', '" & asExamCode & "', '" & asRes1 & "', '" & asRes2 & "', '" & sRefFlag & "','','', '" & sLotNo & "') "
    res = SendQuery(gLocal, SQL)
    If res = -1 Then
        'db_RollBack gLocal
        SaveQuery SQL
        Exit Function
    End If
    
End Function

'WinSock Control ==============================================================================================================
Public Sub WinSock_Listen(argWinSock As Winsock, EquipNum As Integer)
    Dim sWinSockPort As String
    
On Error GoTo IP_Port
    If EquipNum = 1 Then
        sWinSockPort = gSetup.gPort
    Else
        sWinSockPort = gSetup2.gPort
    End If
    
    If sWinSockPort = "0" Or IsNumeric(sWinSockPort) = False Then
        Exit Sub
    End If
    
    If argWinSock.State <> sckClosed Then
        argWinSock.Close
    End If
    
    argWinSock.LocalPort = sWinSockPort
    argWinSock.Listen
    
    If EquipNum = 1 Then
        lblConnect1.Caption = "연결 대기중..."
    Else
        lblConnect2.Caption = "연결 대기중..."
    End If
    
    Exit Sub
IP_Port:
    MsgBox Error(Err.Number) & vbCrLf & "실행중인 프로그램을 확인해 주세요. ", vbCritical
    End

End Sub

Private Sub Winsock1_Close()
    
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    
    Winsock1.LocalPort = gSetup.gPort
    Winsock1.Listen
    
    lblConnect1.Caption = "신호 대기중..."

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then
        Winsock1.Close
    End If
    
    Winsock1.Accept requestID
    lblConnect1.Caption = "연결[" & requestID & "]" & Winsock1.RemoteHostIP
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim sTmp As String
    Dim strSendData
    Dim strResFlag

    Winsock1.GetData sTmp
    Debug.Print sTmp
    If InStr(1, sTmp, "<?xml version") > 0 Then
        gwTmp1 = ""
    End If
    
    gwTmp1 = gwTmp1 & sTmp
    
    Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & " win1 ]" & sTmp
    
    XML_Parsing gwTmp1
    
    Select Case gXML.DataType
    Case "HEL.R"
        strSendData = WinSock_ACK(gXML.Rece_ControlID, "1")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win1 ]" & strSendData
        Winsock1.SendData strSendData
        gwTmp1 = ""
    Case "DST.R"
        strSendData = WinSock_ACK(gXML.Rece_ControlID, "1")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win1 ]" & strSendData
        Winsock1.SendData strSendData
        
        DoSleep 500
        
        strSendData = WinSock_REQ("1")
        
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win1 ]" & strSendData
        Winsock1.SendData strSendData
        gwTmp1 = ""
    Case "OBS.R"
        H232 gXML.Barcode, gXML.DateTime, gXML.EquipCode, gXML.Result, gXML.StatusCode, gXML.InterpretationCode
        strSendData = WinSock_ACK(gXML.Rece_ControlID, "1")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win1 ]" & strSendData
        Winsock1.SendData strSendData
        gwTmp1 = ""
    Case "EOT.R"
        strSendData = WinSock_END(gXML.Rece_ControlID, "1")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win1 ]" & strSendData
        Winsock1.SendData strSendData
        gwTmp1 = ""
    Case "END.R"
        strSendData = WinSock_ACK(gXML.Rece_ControlID, "1")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win1 ]" & strSendData
        Winsock1.SendData strSendData
        gwTmp1 = ""
        
    End Select
    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblConnect1.Caption = "[Error]" & Number & " : " & Description
End Sub

Private Sub Winsock2_Close()
    
    If Winsock2.State <> sckClosed Then
        Winsock2.Close
    End If
    Winsock2.LocalPort = gSetup2.gPort
    Winsock2.Listen
    
    lblConnect2.Caption = "신호 대기중..."

End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
    If Winsock2.State <> sckClosed Then
        Winsock2.Close
    End If
    
    Winsock2.Accept requestID
    lblConnect2.Caption = "연결[" & requestID & "]" & Winsock2.RemoteHostIP
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    'txtBuff2.Text = txtBuff2.Text & sTmp

    Dim sTmp As String
    Dim strSendData
    Dim strResFlag

    Winsock2.GetData sTmp
    
    If InStr(1, sTmp, "<?xml version") > 0 Then
        gwTmp2 = ""
    End If
    
    gwTmp2 = gwTmp2 & sTmp
    
    Save_Raw_Data "[RX" & Format(Time, "hh:mm:ss") & " win2 ]" & sTmp
    
    XML_Parsing1 gwTmp2
    
    Select Case gXML1.DataType1
    Case "HEL.R"
        strSendData = WinSock_ACK1(gXML1.Rece_ControlID1, "2")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win2 ]" & strSendData
        Winsock2.SendData strSendData
        gwTmp2 = ""
    Case "DST.R"
        strSendData = WinSock_ACK1(gXML1.Rece_ControlID1, "2")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win2 ]" & strSendData
        Winsock2.SendData strSendData
        
        DoSleep 500
        
        strSendData = WinSock_REQ1("2")
        
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win2 ]" & strSendData
        Winsock2.SendData strSendData
        gwTmp2 = ""
    Case "OBS.R"
        H232 gXML1.Barcode1, gXML1.DateTime1, gXML1.EquipCode1, gXML1.Result1, gXML1.StatusCode1, gXML1.InterpretationCode1
        strSendData = WinSock_ACK1(gXML1.Rece_ControlID1, "2")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win2 ]" & strSendData
        Winsock2.SendData strSendData
        gwTmp2 = ""
    Case "EOT.R"
        strSendData = WinSock_END1(gXML1.Rece_ControlID1, "2")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win2 ]" & strSendData
        Winsock2.SendData strSendData
        gwTmp2 = ""
    Case "END.R"
        strSendData = WinSock_ACK1(gXML1.Rece_ControlID1, "2")
        Save_Raw_Data "[TX" & Format(Time, "hh:mm:ss") & " win2 ]" & strSendData
        Winsock2.SendData strSendData
        gwTmp2 = ""
        
    End Select
    
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lblConnect2.Caption = "[Error]" & Number & " : " & Description
End Sub


'==============================================================================================================================

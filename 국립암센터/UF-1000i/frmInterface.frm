VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   Caption         =   " UF-1000i Interface"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   15105
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
   ScaleWidth      =   15105
   Begin IF_UF1000i_국립암센터.MDButton cmdSend 
      Height          =   585
      Left            =   12540
      TabIndex        =   66
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "결과전송"
   End
   Begin IF_UF1000i_국립암센터.MDButton cmdReset 
      Height          =   585
      Left            =   11340
      TabIndex        =   65
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "화면정리"
   End
   Begin IF_UF1000i_국립암센터.MDButton cmdCall 
      Height          =   585
      Left            =   10140
      TabIndex        =   64
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "결과조회"
   End
   Begin VB.Frame FrmTempBox 
      Caption         =   "TempBox"
      Height          =   2205
      Left            =   3540
      TabIndex        =   52
      Top             =   6570
      Visible         =   0   'False
      Width           =   9165
      Begin FPSpread.vaSpread vasTemp2 
         Height          =   945
         Left            =   7980
         TabIndex        =   76
         Top             =   420
         Width           =   945
         _Version        =   393216
         _ExtentX        =   1667
         _ExtentY        =   1667
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
      Begin FPSpread.vaSpread vasOrder 
         Height          =   1365
         Left            =   5700
         TabIndex        =   70
         Top             =   330
         Width           =   1905
         _Version        =   393216
         _ExtentX        =   3360
         _ExtentY        =   2408
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
         SpreadDesigner  =   "frmInterface.frx":0790
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
         Left            =   4560
         TabIndex        =   63
         Top             =   210
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton Command14 
         Caption         =   "사용자변경"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1650
         TabIndex        =   59
         Top             =   1650
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdResCall 
         Caption         =   "QC 결과전송"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   58
         Top             =   1650
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Command15"
         Height          =   435
         Left            =   90
         TabIndex        =   57
         Top             =   1050
         Width           =   2325
      End
      Begin VB.CommandButton Command_setup 
         Caption         =   "코드설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2310
         TabIndex        =   56
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command_close 
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   3420
         TabIndex        =   55
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton Command_Config 
         Caption         =   "통신설정"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1200
         TabIndex        =   54
         Top             =   240
         Width           =   1065
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
         Left            =   90
         Style           =   1  '그래픽
         TabIndex        =   53
         Top             =   240
         Value           =   1  '확인
         Width           =   1065
      End
      Begin FPSpread.vaSpread vasIDTmp 
         Height          =   1035
         Left            =   2790
         TabIndex        =   71
         Top             =   840
         Width           =   1095
         _Version        =   393216
         _ExtentX        =   1931
         _ExtentY        =   1826
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
         SpreadDesigner  =   "frmInterface.frx":4D3F
      End
      Begin FPSpread.vaSpread vasWork 
         Height          =   1485
         Left            =   4050
         TabIndex        =   72
         Top             =   480
         Visible         =   0   'False
         Width           =   1905
         _Version        =   393216
         _ExtentX        =   3360
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
         SpreadDesigner  =   "frmInterface.frx":4F42
      End
      Begin FPSpread.vaSpread vasOrderTmp 
         Height          =   1485
         Left            =   6120
         TabIndex        =   73
         Top             =   510
         Width           =   1905
         _Version        =   393216
         _ExtentX        =   3360
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
         SpreadDesigner  =   "frmInterface.frx":5145
      End
   End
   Begin VB.Frame FrmUseControl 
      Caption         =   "UseControl"
      Height          =   975
      Left            =   1470
      TabIndex        =   51
      Top             =   1770
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   930
         Top             =   330
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   150
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         InputLen        =   1
         RThreshold      =   1
         RTSEnable       =   -1  'True
         EOFEnable       =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '아래 맞춤
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   10305
      Width           =   15105
      _ExtentX        =   26644
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
            TextSave        =   "2012-05-17"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 12:08"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Service Center ☎(02)6205-1751"
            TextSave        =   "Service Center ☎(02)6205-1751"
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
   Begin VB.Frame Frame7 
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9975
      Begin Threed.SSPanel SSPanel1 
         Height          =   585
         Left            =   30
         TabIndex        =   60
         Top             =   120
         Width           =   9915
         _Version        =   65536
         _ExtentX        =   17489
         _ExtentY        =   1032
         _StockProps     =   15
         Caption         =   "     UF-1000i INTERFACE"
         ForeColor       =   8388608
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
         BevelOuter      =   0
         Alignment       =   1
         Begin VB.PictureBox Picture1 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   8040
            Picture         =   "frmInterface.frx":5348
            ScaleHeight     =   255
            ScaleWidth      =   285
            TabIndex        =   68
            Top             =   150
            Width           =   315
         End
         Begin VB.TextBox Text_Today 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   6210
            TabIndex        =   61
            Text            =   "2002/02/18"
            Top             =   150
            Width           =   1515
         End
         Begin VB.Label lblUser 
            BackStyle       =   0  '투명
            Caption         =   "사용자"
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
            Height          =   255
            Left            =   8490
            TabIndex        =   69
            Top             =   210
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "검사일자"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   5190
            TabIndex        =   62
            Top             =   210
            Width           =   840
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   9375
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   14925
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   645
         Left            =   4890
         TabIndex        =   75
         Top             =   2610
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox Text4 
         Height          =   1005
         Left            =   1950
         TabIndex        =   74
         Top             =   2280
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   660
         TabIndex        =   47
         Top             =   360
         Width           =   195
      End
      Begin Threed.SSPanel SSPanel3 
         Height          =   285
         Left            =   180
         TabIndex        =   48
         Top             =   390
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   503
         _StockProps     =   15
         Caption         =   "번호"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.76
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   8955
         Left            =   120
         TabIndex        =   49
         Top             =   300
         Width           =   7815
         _Version        =   393216
         _ExtentX        =   13785
         _ExtentY        =   15796
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
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":58D2
         UserResize      =   2
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   8925
         Left            =   8040
         TabIndex        =   50
         Top             =   300
         Width           =   6765
         _Version        =   393216
         _ExtentX        =   11933
         _ExtentY        =   15743
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
         MaxCols         =   8
         ScrollBars      =   2
         SpreadDesigner  =   "frmInterface.frx":9B30
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   4455
      Left            =   60
      TabIndex        =   3
      Top             =   5820
      Visible         =   0   'False
      Width           =   13095
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1575
         Left            =   3720
         TabIndex        =   32
         Top             =   2790
         Width           =   9285
         Begin VB.TextBox txtEquipID 
            Height          =   345
            Left            =   3600
            TabIndex        =   43
            Text            =   "10"
            Top             =   1140
            Width           =   1875
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Rack Pos"
            Height          =   375
            Left            =   7560
            TabIndex        =   42
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command10 
            Caption         =   "결과입력"
            Height          =   375
            Left            =   5880
            TabIndex        =   41
            Top             =   1110
            Width           =   1635
         End
         Begin VB.TextBox txtEquipCode 
            Height          =   345
            Left            =   1710
            TabIndex        =   40
            Text            =   "0ADVI120"
            Top             =   1125
            Width           =   1875
         End
         Begin VB.CommandButton Command9 
            Caption         =   "장비ID조회"
            Height          =   375
            Left            =   60
            TabIndex        =   39
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command8 
            Caption         =   "미검사상세목록"
            Height          =   375
            Left            =   5010
            TabIndex        =   38
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command7 
            Caption         =   "미검사목록"
            Height          =   375
            Left            =   3360
            TabIndex        =   37
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command6 
            Caption         =   "검사상세목록"
            Height          =   375
            Left            =   1710
            TabIndex        =   36
            Top             =   690
            Width           =   1635
         End
         Begin VB.TextBox txtID 
            Height          =   345
            Left            =   6660
            TabIndex        =   35
            Text            =   "05111000003"
            Top             =   720
            Width           =   1875
         End
         Begin VB.CommandButton Command5 
            Caption         =   "검사목록"
            Height          =   375
            Left            =   60
            TabIndex        =   34
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command4 
            Caption         =   "서버시간"
            Height          =   375
            Left            =   60
            TabIndex        =   33
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   1920
            TabIndex        =   45
            Top             =   330
            Width           =   945
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   3150
            TabIndex        =   44
            Top             =   330
            Width           =   945
         End
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   210
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   3360
         Width           =   945
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
         Height          =   615
         Left            =   3450
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   29
         Top             =   1950
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   285
         Left            =   60
         TabIndex        =   28
         Top             =   555
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   240
         TabIndex        =   27
         Top             =   1380
         Width           =   3045
      End
      Begin VB.Frame Frame3 
         Height          =   585
         Left            =   60
         TabIndex        =   20
         Top             =   3780
         Visible         =   0   'False
         Width           =   3675
         Begin VB.TextBox txtEnd 
            Alignment       =   1  '오른쪽 맞춤
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
            Left            =   1950
            TabIndex        =   23
            Top             =   180
            Width           =   885
         End
         Begin VB.TextBox txtStart 
            Alignment       =   1  '오른쪽 맞춤
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
            Left            =   630
            TabIndex        =   22
            Top             =   180
            Width           =   885
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "삭제"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   21
            Top             =   180
            Width           =   705
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "번호"
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
            Left            =   60
            TabIndex        =   25
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " - "
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
            Left            =   1530
            TabIndex        =   24
            Top             =   240
            Width           =   360
         End
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
         Left            =   240
         TabIndex        =   19
         Top             =   1875
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.TextBox txtErr 
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   10260
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   18
         Top             =   1950
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdUp 
         Height          =   525
         Left            =   1260
         Picture         =   "frmInterface.frx":D8E6
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdDown 
         Height          =   525
         Left            =   2010
         Picture         =   "frmInterface.frx":DA15
         Style           =   1  '그래픽
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   285
         Left            =   1710
         TabIndex        =   15
         Top             =   900
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   285
         Left            =   1710
         TabIndex        =   14
         Top             =   570
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   240
         TabIndex        =   12
         Top             =   2850
         Visible         =   0   'False
         Width           =   3045
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
         Height          =   345
         Left            =   5970
         TabIndex        =   11
         Top             =   1500
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   285
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtBarcode 
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
         Left            =   3480
         TabIndex        =   8
         Top             =   1500
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CommandButton cmdWorkList 
         Caption         =   "WorkList 작성"
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
         Height          =   375
         Left            =   30
         TabIndex        =   7
         Top             =   930
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   240
         TabIndex        =   6
         Top             =   2355
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   285
         Left            =   1710
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   1125
         Left            =   10740
         TabIndex        =   5
         Top             =   240
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
         SpreadDesigner  =   "frmInterface.frx":DB47
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   3450
         TabIndex        =   10
         Top             =   240
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
         SpreadDesigner  =   "frmInterface.frx":11FCC
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1125
         Left            =   7110
         TabIndex        =   13
         Top             =   240
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
         SpreadDesigner  =   "frmInterface.frx":121CF
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   5295
         TabIndex        =   26
         Top             =   240
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
         SpreadDesigner  =   "frmInterface.frx":123D2
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   8925
         TabIndex        =   31
         Top             =   240
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
         SpreadDesigner  =   "frmInterface.frx":125D5
      End
      Begin VB.Label lblMT 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "0"
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
         Left            =   9750
         TabIndex        =   46
         Top             =   2370
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin IF_UF1000i_국립암센터.MDButton cmdExit 
      Height          =   585
      Left            =   13740
      TabIndex        =   67
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "종 료"
   End
   Begin VB.Menu MnMain 
      Caption         =   "파일"
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

Const colCheckBox = 1
Const colBarCode = 2
Const colSeqNo = 3
Const colReceno = 4
Const colRack = 5
Const colPos = 6
Const colPID = 7
Const colPName = 8
Const colPSex = 9
Const colPAge = 10
Const colPJumin = 11
Const colState = 12

Const colOrd = 13
Const colRes = 14
Const colDate = 15
Const colTime = 16
Const colTestType = 17

Const colEquipCode = 1
Const colExamCode = 2
Const colExamName = 3
Const colResult = 4
Const colSeq = 5
Const colRCheck = 6

'2004/10/21 이상은
'Const colRefLow = 7
Const colResult1 = 7

Const colRefHigh = 8

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

Dim gMT As String
Dim gComState As Long
Dim gErrState As Long

Dim is_UF_Ord1  As String
Dim is_UF_Ord2  As String
Dim liEquipCode As Long

Dim iResFlag    As Integer

Dim sAllresult  As String
Dim gSelExam        As String

Sub wf_uf1000_res1(asData As String)
'Sample Information Block

    Dim ls_data As String
    
    Dim i, k, m, li_for As Integer
    
    Dim lsID, lsRack, lsPos As String
    Dim lsReview, lsColor, lsClarity, lsFlag, lsTotalCnt As String
    Dim lsFlagRBC, lsFlagWBC, lsFlagEC, lsFlagCAST, lsFlagBACT, lsFlagCond As String
    Dim lsReviewComm, lsTemp As String
    
    Dim iRow As Long
    Dim lsSeqNo As String
    
    Dim lsCha, lsNo As String
    
    Dim lsPriority As String
    
    Dim strTmp As String
    Dim lsBarcode As String
    Dim lsSelExam   As String
    Dim ii As Integer
    
    sAllresult = ""
    
    iResFlag = 0
    
    m = 0
    'gRow = -1
    liEquipCode = 0
    gsBarCode = ""
    
    'ReDim gArrExamRes(1 To 1)
    
    ls_data = asData
    
    lsRack = Trim(Mid(ls_data, 63, 6))
    lsPos = Trim(Mid(ls_data, 69, 2))
    strTmp = mGetP(ls_data, 1, "I")
    lsID = Trim(Mid(ls_data, 71, 15))
    lsBarcode = lsID
    gRow = -1
    
    For iRow = 1 To vasID.DataRowCnt
        If Trim(GetText(vasID, iRow, colBarCode)) = colBarCode Then
            gRow = iRow
            Exit For
        End If
    Next iRow

    If gRow = -1 Then
        gRow = vasID.DataRowCnt + 1
        If vasID.MaxRows < gRow + 1 Then
            vasID.MaxRows = gRow + 1
        End If
    End If
    
    gsBarCode = lsID
    
    
    vasID.SetText colRack, iRow, Trim(lsRack)
    vasID.SetText colPos, iRow, Trim(lsPos)
    
    'vasID.SetText colReceno, iRow, lsID
    vasID.SetText colBarCode, iRow, lsID

    vasID.SetText colState, iRow, "결과"
    
    If Trim(GetText(vasID, iRow, colPID)) = "" Then
        Get_Sample_Info iRow
        
        '**************************************************
        res = Online_XML(gXml_S07, Trim(gsBarCode))
        
        ClearSpread vasTemp


        lsSelExam = ""

        gSelExam = ""

        For ii = 0 To UBound(gExam_Select)
            vasTemp.SetText 1, ii + 1, gExam_Select(ii).TST_CD
            If lsSelExam = "" Then
                lsSelExam = "'" & Trim(GetText(vasTemp, ii + 1, 1)) & "'"
            Else
                lsSelExam = lsSelExam & ",'" & Trim(GetText(vasTemp, ii + 1, 1)) & "'"
            End If
        Next ii

        gSelExam = lsSelExam
        '**************************************************
    
    End If
    
    lsReview = Mid(ls_data, 88, 1)
    If lsReview = "1" Then
        lsReview = "REV"
        SetForeColor vasID, iRow, iRow, 2, vasID.MaxCols, 255, 0, 0
    Else
      lsReview = " "
      SetForeColor vasID, iRow, iRow, 2, vasID.MaxCols, 0, 0, 0
    End If
    Res_Match_Proc "A000", "", lsReview
    lsColor = Mid(ls_data, 106, 1)
    Select Case lsColor
    Case "0"
      lsColor = "None"
    Case "1"
      lsColor = "LyBrown"
    Case "2"
      lsColor = "Yellrow"
    Case "3"
      lsColor = "YBrown"
    Case "4"
      lsColor = "Orange"
    Case "5"
      lsColor = "Red"
    Case "6"
      lsColor = "DBrown"
    Case "7"
      lsColor = "Green"
    Case "8"
      lsColor = "Blue"
    Case "9"
      lsColor = "White"
    Case "*"
      lsColor = "uncertain"
    End Select
    
    Res_Match_Proc "A001", "", lsColor
    
    lsClarity = Mid(ls_data, 107, 1)
    Select Case lsClarity
    Case "0"
      lsClarity = "Clear"
    Case "1"
      lsClarity = "SlHazy"
    Case "2"
      lsClarity = "Hazy"
    Case "3"
      lsClarity = "SlCldy"
    Case "4"
      lsClarity = "Cloudy"
    Case "*"
      lsClarity = "uncertain"
    End Select
    
    Res_Match_Proc "A002", "", lsClarity
    
    lsFlagRBC = Mid(ls_data, 140, 1)
    Select Case lsFlagRBC
        Case "*"
            lsFlagRBC = "Low reliability"
        Case "+"
            lsFlagRBC = "Positivie"
        Case Else
            lsFlagRBC = "Normal"
    End Select
    
    Res_Match_Proc "A003", "", lsFlagRBC
    
    lsFlagWBC = Mid(ls_data, 141, 1)
    Select Case lsFlagWBC
        Case "*"
            lsFlagWBC = "Low reliability"
        Case "+"
            lsFlagWBC = "Positivie"
        Case Else
            lsFlagWBC = "Normal"
    End Select
    
    Res_Match_Proc "A004", "", lsFlagWBC
    
    lsFlagEC = Mid(ls_data, 142, 1)
    Select Case lsFlagEC
        Case "*"
            lsFlagEC = "Low reliability"
        Case "+"
            lsFlagEC = "Positivie"
        Case Else
            lsFlagEC = "Normal"
    End Select
    
    Res_Match_Proc "A005", "", lsFlagEC
    
    lsFlagCAST = Mid(ls_data, 143, 1)
    Select Case lsFlagCAST
        Case "*"
            lsFlagCAST = "Low reliability"
        Case "+"
            lsFlagCAST = "Positivie"
        Case Else
            lsFlagCAST = "Normal"
    End Select
    
    Res_Match_Proc "A006", "", lsFlagCAST
    
    lsFlagBACT = Mid(ls_data, 144, 1)
    Select Case lsFlagBACT
        Case "*"
            lsFlagBACT = "Low reliability"
        Case "+"
            lsFlagBACT = "Positivie"
        Case Else
            lsFlagBACT = "Normal"
    End Select
    
    Res_Match_Proc "A007", "", lsFlagBACT
    
    lsFlagCond = Mid(ls_data, 145, 1)
    Select Case lsFlagCond
        Case "*"
            lsFlagCond = "Low reliability"
        Case "+"
            lsFlagCond = "Positivie"
        Case Else
            lsFlagCond = "Normal"
    End Select
    
    Res_Match_Proc "A008", "", lsFlagCond
    
    lsReviewComm = ""
    lsTemp = Mid(ls_data, 146, 40)
    For i = 1 To 16
        Select Case Mid(lsTemp, i, 1)
            Case "B"
                lsReviewComm = lsReviewComm + "Debris high,"
            Case "C"
                lsReviewComm = lsReviewComm + "RBC/X'TAL,"
            Case "D"
                lsReviewComm = lsReviewComm + "RBC/BACT,"
            Case "E"
                lsReviewComm = lsReviewComm + "RBC/YLC,"
            Case "G"
                lsReviewComm = lsReviewComm + "Urine conductivity abnormal,"
        End Select
    Next
    If Len(lsReviewComm) > 0 Then
        lsReviewComm = Mid(lsReviewComm, 1, Len(lsReviewComm) - 1)
    End If
    
    Res_Match_Proc "A009", "", lsReviewComm

    vasID_Click 2, iRow

End Sub


'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열을 구분자를 이용해 구분해 지정한 위치의 문자열을 구함
'   인수 :
'       1.pText      : 구분자로 구성된 문자열
'       2.pPosiion   : 위치
'       3.pDelimiter : 구분자
'-----------------------------------------------------------------------------'
Public Function mGetP(ByVal pText As String, ByVal pPosition As Integer, _
                      ByVal pDelimiter As String) As String
    
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    Dim i       As Integer

    intPos1 = 0: intPos2 = 0
    
    'pPosition 인수가 1인 경우 For문 Skip
    For i = 1 To pPosition - 1
       intPos1 = intPos2 + 1
       intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
       If intPos2 = 0 Then GoTo ReturnNull
    Next i
    
    '해당 컬럼
    intPos1 = intPos2 + 1
    intPos2 = InStr(intPos2 + 1, pText, pDelimiter)
    If intPos2 = 0 Then intPos2 = Len(pText) + 1
    
    mGetP = Mid$(pText, intPos1, intPos2 - intPos1)
    Exit Function
    
ReturnNull:
    mGetP = ""
End Function

Sub Res_Match_Proc(ByVal asEquipCode As String, ByVal asResult As String, Optional ByVal asEqRes As String = "", Optional ByVal asFlag As String)
            
    Dim i, j, k, iArr, lResRow As Long
    Dim ii As Integer
    Dim iStr As Integer
    Dim iCnt As Integer
    
    Dim lsData As String
    Dim lsTemp As String
    Dim lsSampleType As String
    Dim lsSampleNo As String
    Dim lsSpecimenID As String
    Dim lsOrder As String
    Dim lsID As String
    Dim lsRackID As String
    Dim lsPosNO As String
    Dim lsPriority As String
    
    Dim lsExamCode As String
    Dim lsExamCode1 As String
    Dim lsExamDate As String
    Dim lsEquipCode As String
    Dim lsResult As String
    Dim lsUnit As String
    Dim lsRef As String
    Dim lsState As String
    Dim lsComment As String
    
    Dim iRow As Integer
    
    Dim sCnt As String
    
    Dim lsExamName As String
    Dim lsSeqNo As String
    
    Dim iIndex As Integer
    
    Dim lsGubun As String
    Dim lRow As Long
    
    Dim lsSlip As String
    Dim lsEqRes As String
    
    lsEquipCode = asEquipCode
    lsResult = asResult
    lsEqRes = asEqRes
    
    lRow = gRow
    
    lsID = gsBarCode
    
    lsGubun = "P"

    lResRow = vasRes.DataRowCnt + 1
                
    lsExamCode = ""
    lsExamName = ""
    lsSeqNo = ""
                                            
    If vasRes.MaxRows < lResRow Then vasRes.MaxRows = lResRow

    If Trim(lsResult) <> "" Then iResFlag = iResFlag + 1
            
    k = -1
    For i = LBound(gArrEquip) To UBound(gArrEquip)
        If lsEquipCode = Trim(gArrEquip(i, 2)) Then
            lsExamName = Trim(gArrEquip(i, 4))
            lsSeqNo = Trim(gArrEquip(i, 7))
            
            k = i

            '-- 2012.05.17 수정
'            If Mid(Right(Format(Trim(GetText(vasID, lRow, colReceno)), "0000"), 4), 1, 1) = "3" Then
            If Mid(Right(Format(Trim(GetText(vasID, lRow, colSeqNo)), "0000"), 4), 1, 1) = "3" Then
                lsSlip = "L80"
            Else
                lsSlip = "L61"
            End If
            
            SQL = " Select ExamCode From EquipExam " & CR & _
                  "  Where Equipno = '" & gEquip & "' " & CR & _
                  "    And EquipCode = '" & Trim(lsEquipCode) & "' " & vbCrLf & _
                  "    And examcode like '" & lsSlip & "%' "
            res = db_select_Col(gLocal, SQL)
            lsExamCode = Trim(gReadBuf(0))
            Save_Raw_Data Trim(gReadBuf(0)) & "(" & res & ")" & SQL & vbCrLf & "exmcode : " & lsExamCode
            
            If k > 0 Then Exit For
            '******************************************************************
        End If
    Next i
    
    If k > 0 Then
        vasRes.SetText colEquipCode, lResRow, lsEquipCode
        vasRes.SetText colResult1, lResRow, lsEqRes
        vasRes.SetText colBarCode, lResRow, lsID
    
        vasRes.SetText colExamCode, lResRow, lsExamCode
        vasRes.SetText colExamName, lResRow, lsExamName 'Trim(gArrEquip(k, 3))
        'vasRes.SetText colSeqNo, lResRow, Trim(gArrEquip(k, 6))
        
        If lsExamCode = "L61308" Or lsExamCode = "L80168" Then
            If lsResult <> "0" And lsResult <> "" Then
                vasRes.SetText colResult, lResRow, "REVIEW"
            Else
                vasRes.SetText colResult, lResRow, lsResult
            End If
        Else
            vasRes.SetText colResult, lResRow, lsResult
        End If
        
        vasRes.SetText colResult1, lResRow, lsEqRes
    Else
        If lsExamName <> "" Then
            vasRes.SetText colEquipCode, lResRow, lsEquipCode
            vasRes.SetText colResult1, lResRow, lsEqRes
            vasRes.SetText colBarCode, lResRow, lsID
        
            vasRes.SetText colExamName, lResRow, lsExamName
            If lsSeqNo = "" Then lsSeqNo = lsEquipCode
            vasRes.SetText colSeqNo, lResRow, lsSeqNo
            vasRes.SetText colResult, lResRow, lsResult
            vasRes.SetText colResult1, lResRow, lsEqRes
        End If
    End If
                            
    If Trim(lsExamCode) <> "" And Trim(lsResult) <> "" Then
        sAllresult = sAllresult & lsExamCode & chrTAB & lsResult & chrLF
    End If
    
    If lsExamName <> "" Then Save_Local_One_1 lRow, lResRow, "B"

End Sub

Sub wf_uf1000_res2(asData As String)
'Particle Count Block 1

Dim ls_data As String

Dim li_for As Integer

Dim lsID, lsCnt, lsEquipCode, lsExamCode, lsExamName, lsSeqNo, lsRes1, lsRes, lsEqRes As String
Dim i, j, k, n As Integer
     
ls_data = asData

li_for = gRow

lsID = gsBarCode

lsCnt = Mid(ls_data, 48, 2)
ls_data = Mid(ls_data, 50)


If IsNumeric(lsCnt) Then
    For j = 1 To CInt(lsCnt)
        lsEquipCode = Left(ls_data, 4)
        lsRes = Mid(ls_data, 5, 8)
        lsEqRes = ""
        lsRes1 = ""
        
        ls_data = Mid(ls_data, 13)
        If IsNumeric(lsRes) Then
             lsRes = CStr(CCur(lsRes))
        Else
             'lsRes = ""
        End If
        
    'If Mid(gsBarCode, 7, 1) = "8" Then 'QC Sample
    '    lsEqRes = lsRes
    '    lsRes1 = lsRes
    'Else
        Select Case lsEquipCode
        Case "0000"     'Cast
        If IsNumeric(lsRes) Then
             lsRes = Format(CCur(lsRes), "#0.00")
             lsEqRes = lsRes
             
             lsRes = Format(CCur(lsRes) * 2.9, "#0") '& "/LPF"
             lsRes1 = lsRes
             
            'GRADE for CAST
'            #/uL        /LPF    rank    최종작업
'            1.0            3       1       <1
'            3.8            11      2       1-2
'            6.9            20      3       3-5
'            10.3           30      4       6-10
'            17.2           50      5       11-20
'            17.2           50      6       >20

'            If CCur(lsRes) < 3 Then
'                 lsRes = "Hyaline cast <1"
'            ElseIf CCur(lsRes) < 11 Then
'                 lsRes = "Hyaline cast 1 - 2"
'            ElseIf CCur(lsRes) < 20 Then
'                 lsRes = "Hyaline cast 3 - 5"
'            ElseIf CCur(lsRes) < 30 Then
'                 lsRes = "Hyaline cast 6 - 10"
'            ElseIf CCur(lsRes) < 50 Then
'                 lsRes = "Hyaline cast 11 - 20"
'            Else
'                 lsRes = "Hyaline cast >21"
'            End If
            If CCur(lsRes) < 3 Then
                 lsRes = ""
            ElseIf CCur(lsRes) < 11 Then
                 lsRes = "Hyaline cast <1"
            ElseIf CCur(lsRes) < 20 Then
                 lsRes = "Hyaline cast 1 - 2"
            ElseIf CCur(lsRes) < 30 Then
                 lsRes = "Hyaline cast 3 - 5"
            ElseIf CCur(lsRes) < 50 Then
                 lsRes = "Hyaline cast 6 - 10"
            Else
                 lsRes = "Hyaline cast 11 - 20"
            End If
            
            
'             lsRes = lsRes & " " & lsEqRes
        End If
        Case "0201" 'RBC
        If IsNumeric(lsRes) Then
             lsRes = Format(CCur(lsRes), "#0.0")
             lsEqRes = lsRes
             
'             lsRes = Format(CCur(lsRes) * 0.18, "#0") & "/HPF"
'             lsRes1 = lsRes
             
             lsRes = Format(CCur(lsRes) * 0.18, "#0")
             lsRes1 = lsRes
             
            'GRADE for RBC
            'Negative grade 1 - 2
            'Positive grade 2 - 8
            '    #/ul    (/HPF)  rank
            '<=  16.7    (3.01)  1   0-2
            '<=  33.3    (5.99)  2   3-5
            '<=  61.1    (11.00) 3   6-10
            '<=  116.7   (21.01) 4   11-20
            '<=  172.2   (31.00) 5   21-30
            '<=  283.3   (50.99) 6   31-50
            '<=  555.6   (100.01)    7   51-100
            '>   555.6   (100.01)    8   >100
             
'            #/uL        /HPF    rank    최종작업
'            5.6     1   1      <1
'            16.7        3      2       1-2
'            27.8        5      3       3-4
'            55.6        10     4       5-9
'            111.1       20     5       10-19
'            166.7       30     6       20-29
'            277.8       50     7       30-49
'            555.6       100    8       50-99
'            555.6       100    9       >100
             
             
            If CCur(lsRes) < 1 Then
                 lsRes = "<1"
            ElseIf CCur(lsRes) < 3 Then
                 lsRes = "1 - 2"
            ElseIf CCur(lsRes) < 5 Then
                 lsRes = "3 - 4"
            ElseIf CCur(lsRes) < 10 Then
                 lsRes = "5 - 9"
            ElseIf CCur(lsRes) < 20 Then
                 lsRes = "10 - 19"
            ElseIf CCur(lsRes) < 30 Then
                 lsRes = "20 - 29"
            ElseIf CCur(lsRes) < 50 Then
                 lsRes = "30 - 49"
            ElseIf CCur(lsRes) < 100 Then
                 lsRes = "50 - 99"
            Else
                 lsRes = "≥100"
            End If
            
            lsRes1 = lsRes
            
             'lsRes = lsRes & " " & lsEqRes
        End If
        Case "0202"     'WBC
        If IsNumeric(lsRes) Then
             lsRes = Format(CCur(lsRes), "#0.0")
             lsEqRes = lsRes
             
'             lsRes = Format(CCur(lsRes) * 0.18, "#0") & "/HPF"
'             lsRes1 = lsRes
             
             lsRes = Format(CCur(lsRes) * 0.18, "#0")
             lsRes1 = lsRes
             
            'GRADE for WBC
            'Negative grade 1 - 1
            'Positive grade 2 - 8
            '    #/ul    (/HPF)  rank
            '<=  33.3    (5.99)  1   0-5
            '<=  61.1    (11.00) 2   6-10
            '<=  116.7   (21.01) 3   11-20
            '<=  172.2   (31.00) 4   21-30
            '<=  283.3   (50.99) 5   31-50
            '<=  555.6   (100.01)    6   51-100
            '>   555.6   (100.01)    7   >100
             
'            #/uL        /HPF    rank    최종작업
'            5.6     1   1      <1
'            16.7        3      2       1-2
'            27.8        5      3       3-4
'            55.6        10     4       5-9
'            111.1       20     5       10-19
'            166.7       30     6       20-29
'            277.8       50     7       30-49
'            555.6       100    8       50-99
'            555.6       100    9       >100
             
            If CCur(lsRes) < 1 Then
                 lsRes = "<1"
            ElseIf CCur(lsRes) < 3 Then
                 lsRes = "1 - 2"
            ElseIf CCur(lsRes) < 5 Then
                 lsRes = "3 - 4"
            ElseIf CCur(lsRes) < 10 Then
                 lsRes = "5 - 9"
            ElseIf CCur(lsRes) < 20 Then
                 lsRes = "10 - 19"
            ElseIf CCur(lsRes) < 30 Then
                 lsRes = "20 - 29"
            ElseIf CCur(lsRes) < 50 Then
                 lsRes = "30 - 49"
            ElseIf CCur(lsRes) < 100 Then
                 lsRes = "50 - 99"
            Else
                 lsRes = "≥100"
            End If
            
            lsRes1 = lsRes
            
             'lsRes = lsRes & " " & lsEqRes
        End If
        Case "0100"   'Epi.cell
        If IsNumeric(lsRes) Then
             lsRes = Format(CCur(lsRes), "#0.0")
             lsEqRes = lsRes
             
'             lsRes = Format(CCur(lsRes) * 0.18, "#0") & "/HPF"
'             lsRes1 = lsRes
             
             lsRes = Format(CCur(lsRes) * 0.18, "#0")
             lsRes1 = lsRes
             
            'GRADE for WBC
            'Negative grade 1 - 1
            'Positive grade 2 - 8
            '    #/ul    (/HPF)  rank
            '<=  33.3    (5.99)  1   0-5
            '<=  61.1    (11.00) 2   6-10
            '<=  116.7   (21.01) 3   11-20
            '<=  172.2   (31.00) 4   21-30
            '<=  283.3   (50.99) 5   31-50
            '<=  555.6   (100.01)    6   51-100
            '>   555.6   (100.01)    7   >100
             
'            #/uL        /HPF    rank    최종작업
'            5.6     1   1      <1
'            16.7        3      2       1-2
'            27.8        5      3       3-4
'            55.6        10     4       5-9
'            111.1       20     5       10-19
'            166.7       30     6       20-29
'            277.8       50     7       30-49
'            555.6       100    8       50-99
'            555.6       100    9       >100
             
             
            If CCur(lsRes) <= 1 Then
                 lsRes = "<1"
            ElseIf CCur(lsRes) < 3 Then
                 lsRes = "1 - 2"
            ElseIf CCur(lsRes) < 5 Then
                 lsRes = "3 - 4"
            ElseIf CCur(lsRes) < 10 Then
                 lsRes = "5 - 9"
            ElseIf CCur(lsRes) < 20 Then
                 lsRes = "10 - 19"
            ElseIf CCur(lsRes) < 30 Then
                 lsRes = "20 - 29"
            ElseIf CCur(lsRes) < 50 Then
                 lsRes = "30 - 49"
            ElseIf CCur(lsRes) < 100 Then
                 lsRes = "50 - 99"
            Else
                 lsRes = "≥100"
            End If
             
            lsRes1 = lsRes
            
             'lsRes = lsRes & " " & lsEqRes
        End If
        
        Case "0401" 'Bacteria
        If IsNumeric(lsRes) Then
             lsRes = Format(CCur(lsRes), "#0.0")
             lsEqRes = lsRes

'             lsRes = Format(CCur(lsRes) * 0.18, "#0") & "/HPF"
'             lsRes1 = lsRes
             
             lsRes = Format(CCur(lsRes) * 0.18, "#0")
             lsRes1 = lsRes
             
            'GRADE for BACT
            'Negative grade 1 - 3
            'Positive grade 4 - 4
            '    #/ul    (/HPF)  rank
            '<=  100 (18)    1
            '<=  1000    (180)   2   2. a few
            '<=  10000   (1,800) 3   3. moderate
            '>   10000   (1,800) 4   4. many
             
'                #/uL        /HPF    rank    최종작업
'            <   100        18          1       공란
'            <   1000        180        2       1+
'            <   10000       1800       3       2+
'            >=  10000       1800       4       3+
             
             If CCur(lsRes) < 18 Then
                  lsRes = ""
             ElseIf CCur(lsRes) < 180 Then
                  lsRes = "1+"
             ElseIf CCur(lsRes) < 1800 Then
                  lsRes = "2+"
             Else
                  lsRes = "3+"
             End If
             
             lsRes1 = lsRes
             
             'lsRes = lsRes & " " & lsEqRes
        Else
             If Trim(lsRes) = "*****.**" Then
                lsEqRes = lsRes
                lsRes1 = lsRes
                lsRes = "many"
             End If
        End If
        Case Else
        End Select
        
        Res_Match_Proc lsEquipCode, lsRes, lsEqRes
        
    'End If

    Next j
End If


End Sub

Sub wf_uf1000_res3(asData As String)
'Comment Block1

    Dim ls_data As String
    
    Dim li_for As Integer
    
    Dim lsID, lsCnt, lsEquipCode, lsExamCode, lsExamName, lsSeqNo, lsRes1, lsRes, lsEqRes As String
    Dim i, j, k, n As Integer
    
    ls_data = asData
    
    li_for = gRow
    
    lsCnt = Mid(ls_data, 48, 2)
    ls_data = Mid(ls_data, 50)
    
    lsID = gsBarCode
    
    If IsNumeric(lsCnt) Then
      For j = 1 To CInt(lsCnt)
          lsEquipCode = Left(ls_data, 4) & "F"
          ls_data = Mid(ls_data, 5)
    
          Res_Match_Proc lsEquipCode, "+", "+", "+"
      Next j
    End If

End Sub

Sub wf_uf1000_res4(asData As String)
'Particle Count Block 2
Dim ls_data As String

Dim li_for As Integer

Dim lsID, lsCnt, lsEquipCode, lsExamCode, lsExamName, lsSeqNo, lsRes1, lsRes, lsEqRes As String
Dim i, j, k, n As Integer
     
ls_data = asData

li_for = gRow

lsCnt = Mid(ls_data, 48, 2)
ls_data = Mid(ls_data, 50)

lsID = gsBarCode

If IsNumeric(lsCnt) Then
    For j = 1 To CInt(lsCnt)
        lsEquipCode = Left(ls_data, 4)
        lsRes = Mid(ls_data, 5, 8)
        lsEqRes = ""
        lsRes1 = ""
        
        ls_data = Mid(ls_data, 13)
        If IsNumeric(lsRes) Then
            lsRes = CStr(CCur(lsRes))
        Else
            lsRes = " "
        End If
        
    If Mid(gsBarCode, 7, 1) = "8" Then 'QC Sample
        lsEqRes = lsRes
        lsRes1 = lsRes
    Else
        Select Case lsEquipCode
        Case "0107", "0501" 'SRC, SPERM
             If IsNumeric(lsRes) Then
                 lsRes = Format(CCur(lsRes), "#0")
                 'lsEqRes = lsRes + "/㎕"
                 lsEqRes = lsRes
                 
                 
                 SQL = "select equipres FROM pat_res " & vbCrLf & _
                       "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                       "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                       "  AND equipcode = '" & lsEquipCode & "F'" & vbCrLf & _
                       "  AND barcode = '" & lsID & "' "
                 res = db_select_Col(gLocal, SQL)
'                 If Trim(gReadBuf(0)) <> "+" Then
'                     lsRes = "- (" & lsEqRes & "/㎕)"
'
'                 Else
'                     lsRes = "+ (" & lsEqRes & "/㎕)"
'                     'lsRes = "+ (" & lsRes & ")"
'                 End If

                 If Trim(gReadBuf(0)) <> "+" Then
                     lsRes = "-"
                 Else
                     lsRes = "+"
                 End If
             Else
                 lsRes = " "
             End If
             
             'lsRes = " "
             lsRes1 = lsRes
        Case "0300", "00D9" 'X'TAL, Path.CAST
             If IsNumeric(lsRes) Then
                 lsRes = Format(CCur(lsRes), "#0")
                 lsEqRes = lsRes + "/㎕"
                 lsEqRes = lsRes
                 
                 SQL = "select equipres FROM pat_res " & vbCrLf & _
                       "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                       "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                       "  AND equipcode = '" & lsEquipCode & "F'" & vbCrLf & _
                       "  AND barcode = '" & lsID & "' "
                 res = db_select_Col(gLocal, SQL)
'                 If Trim(gReadBuf(0)) <> "+" Then
'                     lsRes = "- (" & lsEqRes & "/㎕)"
'
'                 Else
'                     lsRes = "+ (" & lsEqRes & "/㎕)"
'
'                     SetForeColor vasID, li_for, li_for, 0, 0, 255
'
'                 End If

                 If Trim(gReadBuf(0)) <> "+" Then
                     lsRes = "-"
                 Else
                     lsRes = "+"
                 End If
             Else
                 lsRes = " "
             End If
             
             'lsRes = " "
             lsRes1 = lsRes
        Case "00DA" 'SPERM
             If IsNumeric(lsRes) Then
                 lsRes = Format(CCur(lsRes), "#0")
                 'lsEqRes = lsRes + "/㎕"
                 lsEqRes = lsRes
                 'If CCur(lsRes) = 0 Then lsRes = " "
                 
                 SQL = "select equipres FROM pat_res " & vbCrLf & _
                       "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                       "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                       "  AND equipcode = '" & lsEquipCode & "F'" & vbCrLf & _
                       "  AND barcode = '" & lsID & "' "
'                 res = db_select_Col(gLocal, SQL)
'                 If Trim(gReadBuf(0)) <> "+" Then
'                     lsRes = "- (" & lsEqRes & "/㎕)"
'
'                 Else
'                     lsRes = "+ (" & lsEqRes & "/㎕)"
'                 End If
             
                 res = db_select_Col(gLocal, SQL)
                 If Trim(gReadBuf(0)) <> "+" Then
                     lsRes = "-"
                 
                 Else
                     lsRes = "+"
                 End If
             Else
                 lsRes = " "
             End If
             'lsRes = " "
             lsRes1 = lsRes
'      case "0500" 'OTHER
        Case "0402" 'YLC
             If IsNumeric(lsRes) Then
                 lsRes = Format(CCur(lsRes), "#0")
                 'lsEqRes = lsRes + "/㎕"
                 lsEqRes = lsRes
                 'If CCur(lsRes) = 0 Then lsRes = " "
                 
                 SQL = "select equipres FROM pat_res " & vbCrLf & _
                       "WHERE examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                       "  AND equipno = '" & gEquip & "' " & vbCrLf & _
                       "  AND equipcode = '" & lsEquipCode & "F'" & vbCrLf & _
                       "  AND barcode = '" & lsID & "' "
                 res = db_select_Col(gLocal, SQL)
                 If Trim(gReadBuf(0)) <> "+" Then
                     lsRes = "- (" & lsEqRes & "/㎕)"
                 
                 Else
                     lsRes = "+ (" & lsEqRes & "/㎕)"
                 End If

             Else
                 lsRes = " "
                 lsEqRes = ""
             End If

             lsRes1 = lsRes
        Case "0502" 'Cond.
             If IsNumeric(lsRes) Then
                 lsRes = Format(CCur(lsRes), "#0")
                 'lsEqRes = lsRes + "/㎕"
                 lsEqRes = lsRes
                 If CCur(lsRes) = 0 Then lsRes = " "
                 
             Else
                 lsRes = " "
             End If
             
             'lsRes = " "
             lsRes1 = lsRes
        Case Else
        End Select
        
        lsRes = ""
        Res_Match_Proc lsEquipCode, lsRes, lsEqRes
        
    End If
        
    Next j
End If


End Sub

Sub wf_uf1000_res5(asData As String)
    'Particle Count Block 2
    Dim ls_data As String
    
    Dim lsID, lsCnt, lsEquipCode, lsExamCode, lsExamName, lsSeqNo, lsRes1, lsRes, lsEqRes As String
    Dim i, j, k, n As Integer
    
    Dim ls_examdate, ls_examtime, ls_result, ls_rcheck As String
    Dim ls_pid, ls_rdate, ls_rno, ls_jumin, ls_sex As String
    Dim lsRBCInfo, lsCondInfo, lsUTIInfo As String
    Dim lRow As Integer
    Dim li_check As Integer
    Dim ld_date As Date
    
    Dim liRet
    
    'gs_age_sexs str_age_sexs
         
    ls_data = asData
    
    lRow = gRow
    lsID = gsBarCode
    
    lsCnt = Mid(ls_data, 48, 2)
    ls_data = Mid(ls_data, 50)
    
    
    If IsNumeric(lsCnt) Then
        For j = 1 To CInt(lsCnt)
            lsEquipCode = Left(ls_data, 4)
            lsRes = Mid(ls_data, 5, 8)
            lsEqRes = ""
            lsRes1 = ""
            
            ls_data = Mid(ls_data, 13)
            If IsNumeric(lsRes) Then
                lsRes = CStr(CInt(lsRes))
            Else
                lsRes = ""
            End If
    
            Select Case lsEquipCode
            Case "0C00" 'RBC-Info : RBC morphological information
                If lsRes = "0" Then
                    lsRes = " "
                ElseIf lsRes = "1" Then
                    lsRes = "Isomorphic?"
                ElseIf lsRes = "2" Then
                    lsRes = "Dymorphic?"
                ElseIf lsRes = "3" Then
                    lsRes = "Mixed?"
                End If
                lsRBCInfo = lsRes
            Case "0C01" 'Cond.-Info : Urine Condensation Information
                 If lsRes = "0" Then
                    lsRes = " "
                ElseIf lsRes = "1" Then
                    lsRes = "RANK1"
                ElseIf lsRes = "2" Then
                    lsRes = "RANK2"
                ElseIf lsRes = "3" Then
                    lsRes = "RANK3"
                ElseIf lsRes = "4" Then
                    lsRes = "RANK4"
                ElseIf lsRes = "5" Then
                    lsRes = "RANK5"
                End If
                lsCondInfo = lsRes
            Case "0C02" 'UTI Information
                 If lsRes = "0" Then
                    lsRes = " "
                ElseIf lsRes = "1" Then
                    lsRes = "UTI?"
                End If
                lsUTIInfo = lsRes
            Case Else
            End Select
            
            Res_Match_Proc lsEquipCode, "", lsRes
            
        Next j
    End If
    
    SetText vasID, "결과", lRow, colState

    'SetText vasID, vasRes.DataRowCnt, lRow, colRCnt
    
    If chkMode.Value = 1 And lRow >= 1 Then
        'res = ToServer(iRow)
        res = Insert_Data(lRow)  '서버에 데이타 전송
        If res = 1 Then
            SetText vasID, "완료", lRow, colState
            SetBackColor vasID, lRow, lRow, colBarCode, colState, 202, 255, 112
        Else
            vasID.Row = lRow
            vasID.Col = colCheckBox
            vasID.Value = 1
            SetText vasID, "실패", lRow, colState
            SetForeColor vasID, lRow, lRow, 1, 1, 255, 0, 0
            SetBackColor vasID, lRow, lRow, colBarCode, colState, 255, 255, 255
        End If
    End If

End Sub

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

Private Sub cmdCall_Click()
    Dim iRow As Long

    ClearSpread vasID
    ClearSpread vasRes
    
    SQL = "select distinct levelname, '', '', '0', '0', examtime, '', '', '', 'F' " & vbCrLf & _
          "from qc_res " & vbCrLf & _
          "where equipno  = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' "
    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
    
    SQL = "select barcode, seqno, receno, diskno, posno, pid, pname, psex, page, jumin, sendflag, count(*), count(*), max(recedate)" & _
          " from pat_res " & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag in ('B','C') " & vbCrLf & _
          "group by diskno, posno, barcode, seqno, receno, pid, pname, psex, page, jumin, sendflag "
    SQL = SQL & vbCrLf & " Union " & vbCrLf
    SQL = SQL & vbCrLf & _
          "select barcode, seqno, receno, diskno, posno, pid, pname, psex, page, jumin, sendflag, count(*), '0',  max(recedate)" & _
          " from pat_res " & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag not in ('B','C') " & vbCrLf & _
          "group by diskno, posno, barcode, seqno, receno,  pid, pname, psex, page, jumin, sendflag " & vbCrLf & _
          "order by diskno,posno"
    res = db_select_Vas(gLocal, SQL, vasID, vasID.DataRowCnt + 1, 2)
    
'    SQL = "select barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, refvalue, panicvalue, max(recedate)" & _
'          "from pat_res " & _
'          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
'          "group by barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, refvalue, panicvalue " & vbCrLf & _
'          "order by diskno,posno"
'    res = db_select_Vas(gLocal, SQL, vasID, 1, 2)
    
    
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    vasSort vasID, colRack, colPos
    
    For iRow = 1 To vasID.DataRowCnt
        Select Case Trim(GetText(vasID, iRow, colState))
         Case "A"
            SetText vasID, "결과", iRow, colState
        Case "B", "C"
            SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasID, "완료", iRow, colState
        Case "O"
            SetText vasID, "오더", iRow, colState
        End Select
    Next iRow
    
    vasID.RowHeight(-1) = 20
End Sub

Private Sub cmdDelete_Click()
    Dim lRow As Long
    Dim lsPID As String
    Dim lsReceNo1 As String
    Dim lsReceNo2 As String
    
    Dim sStart As String
    Dim send As String
    
    sStart = Trim(txtStart.Text)
    send = Trim(txtEnd.Text)
    
    If sStart <> "" And send <> "" Then
        For lRow = sStart To send
            lsPID = Trim(GetText(vasID, lRow, 5))
            lsReceNo1 = Trim(GetText(vasID, lRow, 11))
            lsReceNo2 = Trim(GetText(vasID, lRow, 12))
            SQL = "Delete from pat_res " & vbCrLf & _
                  "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                  "  and equipno = '" & gEquip & "' " & vbCrLf & _
                  "  and pid = '" & lsPID & "' " & vbCrLf & _
                  "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                  "  and receno1 = '" & lsReceNo2 & "' "
            res = SendQuery(gLocal, SQL)
            
            DeleteRow vasID, lRow, lRow
        Next lRow
    Else
        lRow = 1
        Do While lRow <= vasID.DataRowCnt
            vasID.Row = lRow
            vasID.Col = 1
            If vasID.Value = 1 Then
                lsPID = Trim(GetText(vasID, lRow, 5))
                lsReceNo1 = Trim(GetText(vasID, lRow, 11))
                lsReceNo2 = Trim(GetText(vasID, lRow, 12))
                SQL = "Delete from pat_res " & vbCrLf & _
                      "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      "  and equipno = '" & gEquip & "' " & vbCrLf & _
                      "  and pid = '" & lsPID & "' " & vbCrLf & _
                      "  and receno = '" & lsReceNo1 & "' " & vbCrLf & _
                      "  and receno1 = '" & lsReceNo2 & "' "
                res = SendQuery(gLocal, SQL)
                
                DeleteRow vasID, lRow, lRow
            Else
                lRow = lRow + 1
            End If
        Loop
    End If
    
    MsgBox "삭제 완료"
    chkAll.Value = 0
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow + 1
    vasActiveCell vasID, lRow + 1, 2
    vasID_Click 2, lRow + 1
End Sub

Private Sub cmdOrder_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
    
Private Sub cmdQC_Click()
    'frmQCResSch.Show
End Sub

Private Sub cmdResCall_Click()
'    frmResult.Show 0
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer

    Var_Clear
    
    txtData.Text = ""
    txtErr.Text = ""
    
'    If chkAll.Value = 1 Then
'            For i = 1 To vasID.DataRowCnt
'                vasID.Row = i
'                vasID.Col = 1
'
'                If vasID.Value = 1 Then
'                    DeleteRow vasID, i, i
'                    i = i - 1
'                End If
'            Next i
'
'            chkAll.Value = 0
'    Else
'        vasID.Row = 1
'        vasID.Row2 = vasID.MaxRows
'        vasID.Col = 1
'        vasID.Col2 = vasID.MaxCols
'        vasID.BlockMode = True
'        vasID.BackColor = RGB(255, 255, 255)
'        vasID.Action = 3
'        vasID.BlockMode = False
'    End If
    
    SetForeColor vasID, 1, vasID.MaxRows, 1, vasID.MaxCols, 0, 0, 0
    SetForeColor vasRes, 1, vasRes.MaxRows, 1, vasRes.MaxCols, 0, 0, 0
    
    ClearSpread vasID
    ClearSpread vasRes
    
    Text_Today = Format(CDate(Date), "yyyy/mm/dd")
    
    gRow = 0

End Sub

Private Sub cmdResSave_Click()
'    Proc_Result txtBarcode
End Sub

Private Sub cmdSend_Click()
    Dim lRow As Long
    
    For lRow = 1 To vasID.DataRowCnt
        vasID.Row = lRow
        vasID.Col = 1
        If vasID.Value = 1 Then
            res = Insert_Data(lRow)
        
            If res = -1 Then
                SetForeColor vasID, lRow, lRow, 1, colState, 255, 0, 0
                SetText vasID, "실패", lRow, colState
            Else
                vasID.Row = lRow
                vasID.Col = 1
                vasID.Value = 1
                
                SetBackColor vasID, lRow, lRow, 1, colState, 202, 255, 112
                SetText vasID, "완료", lRow, colState
                
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C' " & vbCrLf & _
                      " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                      " And equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, lRow, colBarCode)) & "' "
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

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasID.ActiveRow
    
    vasID.SwapRange 1, lRow, 15, lRow, 1, lRow - 1
    vasActiveCell vasID, lRow - 1, 2
    vasID_Click 2, lRow - 1
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

Private Sub Command1_Click()
'    Hitachi747 Mid(Text2.Text, 2)
End Sub

Private Sub Command10_Click()
'    Dim oerrmsg$
'    Dim ispcid$(), iexamcode$(), iresult$(), ierrflag$(), iequipcd$(), igubun$
'    Dim lRow As Long
'
'    If vasList.DataRowCnt < 1 Then Exit Sub
'
'    ReDim ispcid(vasList.DataRowCnt)
'    ReDim iexamcode(vasList.DataRowCnt)
'    ReDim iresult(vasList.DataRowCnt)
'    ReDim ierrflag(vasList.DataRowCnt)
'    ReDim iequipcd(vasList.DataRowCnt)
'
'    For lRow = 1 To vasList.DataRowCnt
'        ispcid(lRow - 1) = Trim(GetText(vasList, lRow, 1))
'        iexamcode(lRow - 1) = Trim(GetText(vasList, lRow, 6))
'        iresult(lRow - 1) = Trim(GetText(vasList, lRow, 8))
'        ierrflag(lRow - 1) = ""
'        iequipcd(lRow - 1) = Trim(txtEquipCode)
'        'iequipcd(lRow - 1) = ""
'    Next lRow
'    res = sl_online_result_ul_e(oerrmsg, ispcid(), iexamcode(), iresult(), ierrflag(), iequipcd(), "")
'    If res < 0 Then
'        MsgBox "저장 에러"
'    Else
'        MsgBox "저장 확인 : " & res
'    End If

End Sub

Private Sub Command11_Click()
'    Dim oerrmsg$
'    Dim ispcid$(), imach_id$(), ipos_flag$(), irack_id$(), irack_pos$()
'
'    ReDim ispcid(0)
'    ReDim imach_id(0)
'    ReDim ipos_flag(0)
'    ReDim irack_id(0)
'    ReDim irack_pos(0)
'
'    ispcid(0) = Trim(txtID)
'    imach_id(0) = Trim(txtEquipID)
'    ipos_flag(0) = "E"
'    irack_id(0) = "1001"
'    irack_pos(0) = "1"
'
'    res = sl_upd_spc_pos("", ispcid(), imach_id(), ipos_flag(), irack_id(), irack_pos())
'    MsgBox res
End Sub

Private Sub Command12_Click()
'    Dim lsChar As String
'    Dim i As Long
'
'
'    For i = 1 To Len(Text3.Text)
'
'        lsChar = Mid(Text3.Text, i, 1)
'
'        Select Case lsChar
'        Case chrSOH
'            txtData.Text = txtData.Text & lsChar
'            gPreMsg = chrACK
'            MSComm1.Output = chrACK
'            SaveData "[Tx]" & chrACK
'            gACKSig = 1
'            gComState = 0
'
'        Case "["
'            txtData.Text = lsChar
'
'        Case chrLF
'            txtData.Text = txtData.Text & lsChar
'
'            SaveData "[Rx]" & txtData.Text
'
'            LX20 Mid(txtData.Text, 2)
'            gComState = 1
'
'            If gACKSig = 1 Then
'                gPreMsg = chrETX
'                gACKSig = 0
'            Else
'                gPreMsg = chrACK
'                gACKSig = 1
'            End If
'            MSComm1.Output = gPreMsg
'            SaveData "[Tx]" & gPreMsg
'
'            txtData = ""
'        Case chrEOT
'            txtData.Text = lsChar
'
'            If gComState = 1 And vasTemp1.DataRowCnt > 0 Then
'                gPreMsg = chrEOT & chrSOH
'                MSComm1.Output = chrEOT & chrSOH
'                SaveData "[Tx]" & chrEOT & chrSOH
'
'                gComState = 2
'            End If
'        Case chrACK
'            SaveData "[Rx]" & chrACK
'
'            If gComState = 2 Then
'                gOrderMessage = GetText(vasTemp1, 1, 1)
'                DeleteRow vasTemp1, 1, 1
'                gPreMsg = gOrderMessage
'                MSComm1.Output = gOrderMessage
'                SaveData "[Tx]" & gOrderMessage
'                gOrderMessage = ""
'                gComState = 3
'    '        ElseIf gComState = -1 Then
'    '            CX_Init
'            End If
'        Case chrETX
'            SaveData "[Rx]" & chrACK
'
'            gPreMsg = chrEOT
'            MSComm1.Output = chrEOT
'            SaveData "[Tx]" & chrEOT
'
'            If vasTemp1.DataRowCnt > 0 Then
'                gPreMsg = chrEOT & chrSOH
'                MSComm1.Output = chrEOT & chrSOH
'                SaveData "[Tx]" & chrEOT & chrSOH
'
'                gComState = 2
'            Else
'                gComState = 0
'            End If
'
'        Case Else
'            txtData.Text = txtData.Text & lsChar
'        End Select
'    Next
'
'    Text3.Text = ""
End Sub

Private Sub Command13_Click()
    Dim i As Integer
    
    SQL = "select item_code, item_name, m_stype_code, disp_seq, m_item_code from tbl_item"
    res = db_select_Vas(gLocal_1, SQL, vaSpread1)
    
    SQL = "delete from equipexam"
    res = SendQuery(gLocal, SQL)
    
    For i = 1 To vaSpread1.DataRowCnt
    
        SQL = "insert into equipexam(equipno, examcode, equipcode, examname, examtype, seqno, resprec, examflag) " & vbCrLf & _
              "values('C064','" & Trim(GetText(vaSpread1, i, 1)) & "','" & Trim(GetText(vaSpread1, i, 5)) & "','" & Trim(GetText(vaSpread1, i, 2)) & "','" & Trim(GetText(vaSpread1, i, 3)) & "','" & Trim(GetText(vaSpread1, i, 4)) & "','1','1')"
        res = SendQuery(gLocal, SQL)
    Next
    
End Sub

Private Sub Command14_Click()
    frmUserChange.Show 0
    
End Sub

Private Sub Command15_Click()
    
'    Online_XML GetXmlExamCode, "10010700001"

'    vasID.MaxRows = 1
'    SetText vasID, "10010700001", 1, colBarCode
'
'    Get_Sample_Info 1
'
End Sub

Private Sub Command16_Click()

    Call MSComm1_OnComm
    If Mid(Text4.Text, 5, 1) <> "C" Then
          Select Case Mid(Text4.Text, 5, 2)
          Case "01"
               'gRow = -1
               liEquipCode = 0
               gsBarCode = ""
               ReDim gArrExamRes(1 To 1)
               ClearSpread vasRes
               
               Call wf_uf1000_res1(Text4.Text)
          Case "02"
               Call wf_uf1000_res2(Text4.Text)
          Case "03"
               Call wf_uf1000_res3(Text4.Text)
          Case "04"
               Call wf_uf1000_res4(Text4.Text)
          Case "05"
               Call wf_uf1000_res5(Text4.Text)
          End Select
    Else
    'QC
    
    End If
    Text4 = ""
End Sub

Private Sub Command3_Click()
    SQL = "CREATE INDEX resindex1 ON pat_res (examdate,equipno,barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex1 created"
    Else
        MsgBox "resindex1 failed"
    End If
    SQL = "CREATE INDEX resindex2 ON pat_res (examdate,equipno,barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex2 created"
    Else
        MsgBox "resindex2 failed"
    End If
    
    SQL = "CREATE INDEX resindex3 ON pat_res (barcode,examcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex3 created"
    Else
        MsgBox "resindex3 failed"
    End If
    
    SQL = "CREATE INDEX resindex4 ON pat_res (barcode,equipcode)"
    res = SendQuery(gLocal, SQL)
    If res = 1 Then
        MsgBox "resindex4 created"
    Else
        MsgBox "resindex4 failed"
    End If
End Sub

Private Sub Command4_Click()
'    Dim v_date$()
'    Dim v_date_8$()
'    res = sl_sysdate_select(v_date, v_date_8)
'    If res = 1 Then
'        lblDate1.Caption = v_date(0)
'        lblDate2.Caption = v_date_8(0)
'    End If
End Sub

Private Sub Command5_Click()
'    Dim i_spc_no$
'    Dim i_equip_cd$, v_spc_no$(), v_pt_no$(), v_pt_nm$(), v_tst_cd$(), v_tst_nm$()
'    Dim lRow As Long
'    Dim lCol As Long
'
'    ClearSpread vasList
'
'    i_spc_no = Trim(txtID)
'    res = sl_sel_spcno_tstcd_all&(i_spc_no, i_equip_cd, v_spc_no(), v_pt_no(), v_pt_nm(), v_tst_cd(), v_tst_nm())
'    If res > 0 Then
'        For lRow = LBound(v_tst_cd) To res - 1
'            vasList.SetText 1, lRow + 1, i_spc_no
'            vasList.SetText 2, lRow + 1, i_equip_cd
'            vasList.SetText 3, lRow + 1, v_spc_no(lRow)
'            vasList.SetText 4, lRow + 1, v_pt_no(lRow)
'            vasList.SetText 5, lRow + 1, v_pt_nm(lRow)
'            vasList.SetText 6, lRow + 1, v_tst_cd(lRow)
'            vasList.SetText 7, lRow + 1, v_tst_nm(lRow)
'        Next lRow
'    ElseIf res = 0 Then
'        MsgBox "검사 내역이 존재하지 않습니다"
'    End If
    
End Sub

Private Sub Command6_Click()
'    Dim i_spc_no$
'    Dim i_equip_cd$
'    Dim v_spc_no$(), v_pt_no$(), v_pt_nm$(), v_tst_frct_cd$(), v_tst_frct_nm$(), v_acpt_dte$(), v_acpt_no$(), v_sex$(), v_age$(), v_spc_cd$(), v_spc_nm$(), v_tst_cd$(), v_tst_nm$()
'    Dim lRow As Long
'    Dim lCol As Long
'    Dim i As Long
'
'    i_spc_no = Trim(txtID)
'    res = sl_sel_spcno_tstcd_all_sub(i_spc_no, i_equip_cd, v_spc_no(), v_pt_no(), v_pt_nm(), v_tst_frct_cd(), v_tst_frct_nm(), v_acpt_dte(), v_acpt_no(), v_sex(), v_age(), v_spc_cd(), v_spc_nm(), v_tst_cd(), v_tst_nm())
'    If res < 0 Then
'        MsgBox "검사 내역이 존재하지 않습니다"
'    Else
'        i = 0
'        lRow = 1
'        Do While i < UBound(v_spc_no)
'            vasList.SetText 1, lRow, i_spc_no
'            vasList.SetText 2, lRow, i_equip_cd
'            vasList.SetText 3, lRow, v_spc_no(i)
'            vasList.SetText 4, lRow, v_pt_no(i)
'            vasList.SetText 5, lRow, v_pt_nm(i)
'            vasList.SetText 6, lRow, v_tst_frct_cd(i)
'            vasList.SetText 7, lRow, v_tst_frct_nm(i)
'            vasList.SetText 8, lRow, v_acpt_dte(i)
'            vasList.SetText 9, lRow, v_acpt_no(i)
'            vasList.SetText 10, lRow, v_sex(i)
'            vasList.SetText 11, lRow, v_age(i)
'            vasList.SetText 12, lRow, v_spc_cd(i)
'            vasList.SetText 13, lRow, v_spc_nm(i)
'            vasList.SetText 14, lRow, v_tst_cd(i)
'            vasList.SetText 15, lRow, v_tst_nm(i)
'
'            lRow = lRow + 1
'            If lRow > vasList.MaxRows Then vasList.MaxRows = lRow
'            i = i + 1
'        Loop
'    End If
End Sub

Private Sub Command7_Click()
'    Dim i_spc_no$, i_equip_cd$
'    Dim v_spc_no$(), v_pt_no$(), v_pt_nm$(), v_tst_cd$(), v_tst_nm$()
'    Dim lRow As Long
'    Dim lCol As Long
'
'    ClearSpread vasList
'
'    i_spc_no = Trim(txtID)
'    res = sl_sel_spcno_tstcd_unin&(i_spc_no, i_equip_cd, v_spc_no(), v_pt_no(), _
'                                    v_pt_nm(), v_tst_cd(), v_tst_nm())
'    If res > 0 Then
'        For lRow = LBound(v_tst_cd) To UBound(v_tst_cd) - 1
'            vasList.SetText 1, lRow + 1, i_spc_no
'            vasList.SetText 2, lRow + 1, i_equip_cd
'            vasList.SetText 3, lRow + 1, v_spc_no(lRow)
'            vasList.SetText 4, lRow + 1, v_pt_no(lRow)
'            vasList.SetText 5, lRow + 1, v_pt_nm(lRow)
'            vasList.SetText 6, lRow + 1, v_tst_cd(lRow)
'            vasList.SetText 7, lRow + 1, v_tst_nm(lRow)
'        Next lRow
'    ElseIf res = 0 Then
'        MsgBox "검사 내역이 존재하지 않습니다"
'    End If
End Sub

Private Sub Command8_Click()
'    Dim i_spc_no$
'    Dim i_equip_cd$
'    Dim v_spc_no$(), v_pt_no$(), v_pt_nm$(), v_tst_frct_cd$(), v_tst_frct_nm$(), v_acpt_dte$(), v_acpt_no$(), v_sex$(), v_age$(), v_spc_cd$(), v_spc_nm$(), v_tst_cd$(), v_tst_nm$()
'    Dim lRow As Long
'    Dim lCol As Long
'    Dim i As Long
'
'    i_spc_no = Trim(txtID)
'    res = sl_sel_spcno_tstcd_unin_sub(i_spc_no, i_equip_cd, v_spc_no(), v_pt_no(), v_pt_nm(), v_tst_frct_cd(), v_tst_frct_nm(), v_acpt_dte(), v_acpt_no(), v_sex(), v_age(), v_spc_cd(), v_spc_nm(), v_tst_cd(), v_tst_nm())
'    If res < 0 Then
'        MsgBox "검사 내역이 존재하지 않습니다"
'    ElseIf res = 0 Then
'
'    Else
'        i = 0
'        lRow = 1
'        Do While i < UBound(v_spc_no)
'            vasList.SetText 1, lRow, i_spc_no
'            vasList.SetText 2, lRow, i_equip_cd
'            vasList.SetText 3, lRow, v_spc_no(i)
'            vasList.SetText 4, lRow, v_pt_no(i)
'            vasList.SetText 5, lRow, v_pt_nm(i)
'            vasList.SetText 6, lRow, v_tst_frct_cd(i)
'            vasList.SetText 7, lRow, v_tst_frct_nm(i)
'            vasList.SetText 8, lRow, v_acpt_dte(i)
'            vasList.SetText 9, lRow, v_acpt_no(i)
'            vasList.SetText 10, lRow, v_sex(i)
'            vasList.SetText 11, lRow, v_age(i)
'            vasList.SetText 12, lRow, v_spc_cd(i)
'            vasList.SetText 13, lRow, v_spc_nm(i)
'            vasList.SetText 14, lRow, v_tst_cd(i)
'            vasList.SetText 15, lRow, v_tst_nm(i)
'
'            lRow = lRow + 1
'            If lRow > vasList.MaxRows Then vasList.MaxRows = lRow
'            i = i + 1
'        Loop
'    End If

End Sub

Private Sub Command9_Click()
'    Dim i_equip_cd$
'    Dim machine_id$(), equip_cd$(), equip_nm$()
'    Dim lRow As Long
'    Dim lCol As Long
'
'    ClearSpread vasList
'    i_equip_cd = Trim(txtEquipCode)
'    res = sl_sel_machine_id(i_equip_cd, machine_id(), equip_cd(), equip_nm())
'    If res > 0 Then
'        For lRow = LBound(machine_id) To UBound(machine_id)
'            vasList.SetText 1, lRow + 1, i_equip_cd
'            vasList.SetText 2, lRow + 1, machine_id(lRow)
'            vasList.SetText 3, lRow + 1, equip_cd(lRow)
'            vasList.SetText 4, lRow + 1, equip_nm(lRow)
'        Next lRow
'    End If
End Sub

Private Sub Form_Load()
    Dim sDate As String
            
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
'    Me.Height = 11190
'    Me.Width = 15360

    cmdReset_Click
    
    GetSetup
    
    lblUser.Caption = Trim(gIFUser)
    
    
    MSComm1.CommPort = gSetup.gPort
    MSComm1.RTSEnable = gSetup.gRTSEnable
    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
    Text_Today = Format(CDate(GetDateFull), "yyyy/mm/dd")

    GetExamCode
        
    sDate = Format(DateAdd("y", CDate(Text_Today.Text), -30), "yyyymmdd")
    SQL = "delete from pat_res where examdate < '" & sDate & "'"
    res = SendQuery(gLocal, SQL)
    
    '******************************************************
    SQL = " Select cutoff From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column cutoff text(20) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select cutoffflag From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column cutoffflag long "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select negvalue From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column negvalue text(10) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select posvalue From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column posvalue text(10) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select posequal From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column posequal long "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select negequal From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column negequal long "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Select resgubun From equipexam "
    res = db_select_Col(gLocal, SQL)
    If res = -1 Then
        SQL = " Alter Table equipexam Add Column resgubun text(1) "
        res = SendQuery(gLocal, SQL)
    End If
    
    SQL = " Alter Table equipexam Alter column seqno text(3) "
    res = SendQuery(gLocal, SQL)
    '******************************************************
    
End Sub

'-- seq
''Function Get_Sample_Info(ByVal asRow As Long) As Integer
''    Dim lsBarcode   As String
''    Dim lsPID       As String
''    Dim lsReceNo    As String
''    Dim sRes        As String
''    Dim lsExamPart  As String
''
''    Get_Sample_Info = -1
''
''    '샘플 환자 정보 가져오기
''    lsReceNo = Trim(GetText(vasID, asRow, colReceno))
''
''    If Mid(Right(Format(Trim(GetText(vasID, asRow, colReceno)), "0000"), 4), 1, 1) = "3" Then
''        lsExamPart = "L80"
''    Else
''        lsExamPart = "L61"
''    End If
''
'''    If Trim(lsReceNo) = "" Then: Exit Function
''    sRes = Online_Urine(gXml_S18, Format(Text_Today.Text, "YYYYMMDD"), lsReceNo, lsExamPart)
''    'If gUrine_Info_Select(0).SPC_NO <> "" Then
''    If sRes = 1 Then
''        SetText vasID, gUrine_Info_Select(0).SPC_NO, asRow, colBarCode
''        SetText vasID, gUrine_Info_Select(0).PT_NO, asRow, colPID
''        SetText vasID, gUrine_Info_Select(0).PT_NM, asRow, colPName
''        SetText vasID, gUrine_Info_Select(0).SEX, asRow, colPSex
''        SetText vasID, gUrine_Info_Select(0).AGE, asRow, colPAge
''        SetText vasID, gUrine_Info_Select(0).ACPTNO_1, asRow, colReceno
'''        SetText vasID, gPat_Info_Select.ACPT_DTETM, asRow, colDate
''        SetText vasID, Mid(gUrine_Info_Select(0).ACPT_DTETM, 1, 10), asRow, colDate
''        SetText vasID, gUrine_Info_Select(0).SPC_CD_1, asRow, colSeqNo
''
''        vasID.RowHeight(asRow) = 20
''
''        Get_Sample_Info = 1
''    End If
''End Function


'-- 바코드
Function Get_Sample_Info(ByVal asRow As Long) As Integer
Dim lsBarcode As String
Dim lsPID As String
Dim lsReceNo As String
Dim sRes As String

Dim lsReceDate As String
Dim lsExamPart As String

    Get_Sample_Info = -1
    
    '샘플 환자 정보 가져오기
    
    lsBarcode = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
    sRes = Online_XML(gXml_S03, lsBarcode)
'    If sRes = 1 Then
        SetText vasID, gPat_Info_Select.PT_NO, asRow, colPID
        SetText vasID, gPat_Info_Select.PT_NM, asRow, colPName
        SetText vasID, gPat_Info_Select.SEX, asRow, colPSex
        SetText vasID, gPat_Info_Select.AGE, asRow, colPAge
        SetText vasID, gPat_Info_Select.ACPTNO_1, asRow, colSeqNo
        SetText vasID, gPat_Info_Select.ACPT_DTETM, asRow, colDate
        SetText vasID, Mid(gPat_Info_Select.ACPT_DTETM, 1, 10), asRow, colDate
        SetText vasID, gPat_Info_Select.SPC_CD_1, asRow, colReceno

        vasID.RowHeight(asRow) = 20
        
        Get_Sample_Info = 1
'    End If


End Function

Function Get_Sample_Info_Local(ByVal asRow As Long) As Integer
    Dim lsBarcode As String
    Dim lsPID As String
    Dim lsReceNo As String
    Dim sRes As String

    Get_Sample_Info_Local = -1
    
    '샘플 환자 정보 가져오기
    lsBarcode = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
    SQL = " Select pid, pname, psex, page, seqno, recedate, receno From pat_res " & CR & _
          " Where equipno = '" & gEquip & "' " & CR & _
          " And examdate = '" & Format(Text_Today.Text, "YYYYMMDD") & "' "
    res = db_select_Col(gLocal, SQL)
    
    If res = 1 Then
        SetText vasID, Trim(gReadBuf(0)), asRow, colPID
        SetText vasID, Trim(gReadBuf(1)), asRow, colPName
        SetText vasID, Trim(gReadBuf(2)), asRow, colPSex
        SetText vasID, Trim(gReadBuf(3)), asRow, colPAge
        SetText vasID, Trim(gReadBuf(4)), asRow, colSeqNo
        SetText vasID, Trim(gReadBuf(5)), asRow, colDate
        SetText vasID, Trim(gReadBuf(6)), asRow, colReceno
        
        vasID.RowHeight(asRow) = 20
        
        Get_Sample_Info_Local = 1
    End If
End Function

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
    Dim i, j As Long
    
    ClearSpread vasTemp
    GetExamCode = -1
    
    SQL = "Select equipcode, examcode, examname, reflow, refhigh, seqno " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  examcode "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    If res > 0 Then
        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 7)
    Else
        SaveQuery SQL
        Exit Function
    End If
        
    For i = 1 To vasTemp.DataRowCnt
        gArrEquip(i, 1) = i
        For j = 1 To 6
            gArrEquip(i, j + 1) = Trim(GetText(vasTemp, i, j))
        Next j
    Next i
    
    GetExamCode = 1
End Function

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If

'    Call dce_close_env      ' Server와 연결을 끊는 곳
    DisConnect_Local
    
    Unload Me
    
    End
    
End Sub

Private Sub MDButton3_Click()

End Sub

Private Sub MnExamConfig_Click()
    frmOrderCode.Show 1
    GetExamCode
End Sub

Private Sub MnExit_Click()
    Unload Me
End Sub

Private Sub MnTConfig_Click()
    frmConfig.Show 1
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

Private Sub MSComm1_OnComm()
    Dim lsChar As String
    Dim i As Integer
    Dim strExamCd As String
    
    lsChar = MSComm1.Input
    Debug.Print lsChar
    Save_Raw_Data "UF1000i-" & "[LOW]" & lsChar
    
    Select Case lsChar
    Case chrACK
        MSComm1.Output = is_UF_Ord2
    Case chrSTX
        
        txtData.Text = ""
        
        txtData.Text = txtData.Text & lsChar
    Case chrETX
        txtData.Text = txtData.Text & lsChar
        
        Save_Raw_Data "UF1000i-" & "[RX]" & txtData.Text
        
        'UF1000
        txtData.Text = Mid(txtData.Text, 2, Len(txtData.Text) - 2)
        If Mid(txtData.Text, 1, 1) = "R" Then  '오더요구 데이터 입니다
'            gRow = -1
'            liEquipCode = 0
'            gsBarCode = ""
'            ReDim gArrExamRes(1 To 1)
'            ClearSpread vasRes
'
'            Call wf_uf1000_res1(txtData.Text)
            
            gsBarCode = ""
            gsBarCode = Trim(Mid(txtData.Text, 6, 15))
            'gsBarCode = "12051604246"
            
            'gsBarCode = "12051701164"
'            res = Online_XML(gXml_S07, Trim(gsBarCode))
            res = Online_XML(gXml_S07, Trim(gsBarCode))
            'R1441    12345678901  0001" + " 1" + "000000000000000000" + ETX
            
            strExamCd = ""
            For i = 0 To UBound(gExam_Select)
                If i = 0 Then
                    strExamCd = strExamCd & " and (examcode = '" & Trim(gExam_Select(i).TST_CD) & "' "
                Else
                    strExamCd = strExamCd & " or examcode = '" & Trim(gExam_Select(i).TST_CD) & "' "
                End If
            Next
            
            If strExamCd <> "" Then
                strExamCd = strExamCd & ")"
            End If
            
            SQL = " Select ExamCode From EquipExam " & CR & _
                  "  Where Equipno = '" & gEquip & "' " & CR & _
                     strExamCd & vbCrLf '& _
                  "    And examcode like '" & lsSlip & "%' "
            res = db_select_Col(gLocal, SQL)
            'lsExamCode = Trim(gReadBuf(0))

            If res = 1 And Trim(gReadBuf(0)) <> "" Then
'                is_UF_Ord1 = chrSTX + "S144" & "1" + String(Date, "yyyymmdd") + Mid(txtData.Text, 6, 23) + Mid(txtData.Text, 5, 1) + "1" + Space(56) + String(Date, "yyyymmdd") + String(Date, "hh:mm") + "***" + Space(143) + chrETX
'                is_UF_Ord2 = chrSTX + "S244" & "1" + String(Date, "yyyymmdd") + Mid(txtData.Text, 6, 23) + Mid(txtData.Text, 5, 1) + Space(226) + chrETX
                
                             is_UF_Ord1 = chrSTX
                is_UF_Ord1 = is_UF_Ord1 & "S"   'fix
                is_UF_Ord1 = is_UF_Ord1 & "1"   'fix
                is_UF_Ord1 = is_UF_Ord1 & "44"  'UF-1000i
                is_UF_Ord1 = is_UF_Ord1 & "1"
                is_UF_Ord1 = is_UF_Ord1 & Format(Date, "yyyymmdd")
                is_UF_Ord1 = is_UF_Ord1 & Mid(txtData.Text, 6, 23)  'barcode(15),rack(6),pos(4)
                is_UF_Ord1 = is_UF_Ord1 & Mid(txtData.Text, 5, 1)   'inquery mode >>> "1": Real-time, "2": Batch
                is_UF_Ord1 = is_UF_Ord1 & "1"                       'Order "0": Not analyze,    "1": Sediment (SEDch + BACch),  "2": Only Bacteria (BACch)
                is_UF_Ord1 = is_UF_Ord1 & Space(56)                 'Patient ID(16), Sample Comment(40)
                is_UF_Ord1 = is_UF_Ord1 & Format(Date, "yyyymmdd")
                is_UF_Ord1 = is_UF_Ord1 & Format(Date, "hh:mm")
                is_UF_Ord1 = is_UF_Ord1 & "***"                     'Sample Source 1 >>  "0": OP.CLCT, "1": Morning, "2": Timed
                                                                    '                    "3": AF. Meal, "4": Cath, "*": Uncertain
                                                                    'Sample Color 1 >>   "0": None, "1": LyBrown, "2": Yellow,
                                                                    '                    "3": YBrown, "4": Orange, "5": Red,
                                                                    '                    "6": DBrown, "7": Green, "8": Blue,
                                                                    '                    "9": White, "*": uncertain
                                                                    'Sample Clarity 1 >> "0": Clear, "1": SlHazy, "2": Hazy,
                                                                    '                    "3": SlCldy, "4": Cloudy, "*": Uncertain
                is_UF_Ord1 = is_UF_Ord1 & String(143, "0")          'Reserved >> "00 … 00": all zeros, fixed
                is_UF_Ord1 = is_UF_Ord1 & chrETX
                                       
                             is_UF_Ord2 = chrSTX
                is_UF_Ord2 = is_UF_Ord2 & "S"
                is_UF_Ord2 = is_UF_Ord2 & "2"
                is_UF_Ord2 = is_UF_Ord2 & "44"
                is_UF_Ord2 = is_UF_Ord2 & "1"
                is_UF_Ord2 = is_UF_Ord2 & Format(Date, "yyyymmdd")
                is_UF_Ord2 = is_UF_Ord2 & Mid(txtData.Text, 6, 23)
                is_UF_Ord2 = is_UF_Ord2 & Mid(txtData.Text, 5, 1)
                is_UF_Ord2 = is_UF_Ord2 & Space(205)
                is_UF_Ord2 = is_UF_Ord2 & String(11, "0")
                is_UF_Ord2 = is_UF_Ord2 & chrETX
                                      
            Else
                '-- 오더없음
'                is_UF_Ord1 = chrSTX + "S144" & "0" + String(Date, "yyyymmdd") + Mid(txtData.Text, 6, 23) + Mid(txtData.Text, 5, 1) + "0" + Space(56) + String(Date, "yyyymmdd") + String(Date, "hh:mm") + "***" + Space(143) + chrETX
'                is_UF_Ord2 = chrSTX + "S244" & "0" + String(Date, "yyyymmdd") + Mid(txtData.Text, 6, 23) + Mid(txtData.Text, 5, 1) + Space(226) + chrETX
                
'                is_UF_Ord1 = chrSTX + "S144" & "0" + Format(Date, "yyyymmdd") + _
                                       Mid(txtData.Text, 6, 23) + _
                                       Mid(txtData.Text, 5, 1) + _
                                       "0" + _
                                       Space(56) + _
                                       Format(Date, "yyyymmdd") + _
                                       Format(Date, "hh:mm") + _
                                       "***" + _
                                       String(143, "0") + _
                                       chrETX
                                       
'                is_UF_Ord2 = chrSTX + "S244" & "0" + Format(Date, "yyyymmdd") + _
                                      Mid(txtData.Text, 6, 23) + _
                                      Mid(txtData.Text, 5, 1) + _
                                      Space(205) + _
                                      String(11, "0") + _
                                      chrETX
                                      
            

                             is_UF_Ord1 = chrSTX
                is_UF_Ord1 = is_UF_Ord1 & "S"   'fix
                is_UF_Ord1 = is_UF_Ord1 & "1"   'fix
                is_UF_Ord1 = is_UF_Ord1 & "44"  'UF-1000i
                is_UF_Ord1 = is_UF_Ord1 & "1"
                is_UF_Ord1 = is_UF_Ord1 & Format(Date, "yyyymmdd")
                is_UF_Ord1 = is_UF_Ord1 & Mid(txtData.Text, 6, 23)  'barcode(15),rack(6),pos(4)
                is_UF_Ord1 = is_UF_Ord1 & Mid(txtData.Text, 5, 1)   'inquery mode >>> "1": Real-time, "2": Batch
                is_UF_Ord1 = is_UF_Ord1 & "0"                       'Order "0": Not analyze,    "1": Sediment (SEDch + BACch),  "2": Only Bacteria (BACch)
                is_UF_Ord1 = is_UF_Ord1 & Space(56)                 'Patient ID(16), Sample Comment(40)
                is_UF_Ord1 = is_UF_Ord1 & Format(Date, "yyyymmdd")
                is_UF_Ord1 = is_UF_Ord1 & Format(Date, "hh:mm")
                is_UF_Ord1 = is_UF_Ord1 & "***"                     'Sample Source 1 >>  "0": OP.CLCT, "1": Morning, "2": Timed
                                                                    '                    "3": AF. Meal, "4": Cath, "*": Uncertain
                                                                    'Sample Color 1 >>   "0": None, "1": LyBrown, "2": Yellow,
                                                                    '                    "3": YBrown, "4": Orange, "5": Red,
                                                                    '                    "6": DBrown, "7": Green, "8": Blue,
                                                                    '                    "9": White, "*": uncertain
                                                                    'Sample Clarity 1 >> "0": Clear, "1": SlHazy, "2": Hazy,
                                                                    '                    "3": SlCldy, "4": Cloudy, "*": Uncertain
                is_UF_Ord1 = is_UF_Ord1 & String(143, "0")          'Reserved >> "00 … 00": all zeros, fixed
                is_UF_Ord1 = is_UF_Ord1 & chrETX
                                       
                             is_UF_Ord2 = chrSTX
                is_UF_Ord2 = is_UF_Ord2 & "S"
                is_UF_Ord2 = is_UF_Ord2 & "2"
                is_UF_Ord2 = is_UF_Ord2 & "44"
                is_UF_Ord2 = is_UF_Ord2 & "1"
                is_UF_Ord2 = is_UF_Ord2 & Format(Date, "yyyymmdd")
                is_UF_Ord2 = is_UF_Ord2 & Mid(txtData.Text, 6, 23)
                is_UF_Ord2 = is_UF_Ord2 & Mid(txtData.Text, 5, 1)
                is_UF_Ord2 = is_UF_Ord2 & Space(205)
                is_UF_Ord2 = is_UF_Ord2 & String(11, "0")
                is_UF_Ord2 = is_UF_Ord2 & chrETX
                                                  
            
            
            End If
            
            MSComm1.Output = chrACK
            Save_Raw_Data "UF1000i-" & "[TX]" & chrACK
            
            MSComm1.Output = is_UF_Ord1
            Save_Raw_Data "UF1000i-" & "[TX]" & is_UF_Ord1
            
        ElseIf Mid(txtData.Text, 1, 1) = "D" Then  '결과
            If Mid(txtData.Text, 5, 1) <> "C" Then
                  Select Case Mid(txtData.Text, 5, 2)
                  Case "01"
                       gRow = -1
                       liEquipCode = 0
                       gsBarCode = ""
                       ReDim gArrExamRes(1 To 1)
                       ClearSpread vasRes
                       
                       Call wf_uf1000_res1(txtData.Text)
                  Case "02"
                       Call wf_uf1000_res2(txtData.Text)
                  Case "03"
                       Call wf_uf1000_res3(txtData.Text)
                  Case "04"
                       Call wf_uf1000_res4(txtData.Text)
                  Case "05"
                       Call wf_uf1000_res5(txtData.Text)
                  End Select
            Else
            'QC
            
            End If
              
            MSComm1.Output = chrACK
            Save_Raw_Data "UF1000i-" & "[TX]" & chrACK
        End If
        'Proc_Result
        
    Case Else
        txtData.Text = txtData.Text & lsChar
    End Select

End Sub

Sub SendOrder()
    Dim sSendOrder As String
    
    If Len(gOrderMessage) > 240 Then
        
        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If
        
        sSendOrder = CStr(gOrderCnt) & Left(gOrderMessage, 240) & chrETB
        gOrderMessage = Mid(gOrderMessage, 241)
        
        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
        SaveQuery sSendOrder, 1
        
        gOrderCnt = gOrderCnt + 1
        comSend = "stENQ"
        
        gPreMsg = sSendOrder
        
        Save_Raw_Data "[Tx]" & gPreMsg
        MSComm1.Output = sSendOrder
        
    Else
        If gOrderCnt = 8 Then
            gOrderCnt = 0
        End If
        
        sSendOrder = CStr(gOrderCnt) & gOrderMessage & chrETX
        sSendOrder = chrSTX & sSendOrder & CheckSum(sSendOrder) & chrCR & chrLF
                
        gOrderMessage = ""
        comSend = "stOrder"
        
        gPreMsg = sSendOrder
        
        Save_Raw_Data "[Tx]" & gPreMsg
        MSComm1.Output = sSendOrder
    End If
End Sub

Private Sub Modular(asData As String)
    Dim i           As Integer
    Dim iIndex      As Integer
    
    Dim lsData      As String
    Dim lsTemp      As String
    
    Dim lsHead      As String
    Dim lsPatient   As String
    Dim lsRequest   As String
    Dim lsOrder     As String

    Dim lsMessage   As String
    
    Dim lsMSGflag   As String
    
    lsMessage = ""
    
    If asData = "" Then
        Exit Sub
    End If
    
    ClearSpread vasRes
    ClearSpread vasResTemp
    
    iIndex = 0
    lsData = asData
    
    i = InStr(1, lsData, Chr(13))
    Do While i > 0
        lsTemp = Mid(lsData, 1, i - 1)
        lsData = Mid(lsData, i + 1)
        
        Select Case Left(lsTemp, 1)
        Case "H"
            lsHead = lsTemp
        Case "P"
            lsPatient = lsTemp
        Case "O"
            lsOrder = lsTemp
        Case "Q"
            lsRequest = lsTemp
            lsMSGflag = "Q"
        Case "R"
            iIndex = iIndex + 1
            If iIndex > vasResTemp.MaxRows Then vasResTemp.MaxRows = iIndex

            SetText vasResTemp, lsTemp, iIndex, 1
            
            lsMSGflag = "R"
        Case "C"
            SetText vasRes, lsTemp, iIndex, 2
        Case "L"
            lsMessage = lsTemp
        End Select
        
        i = InStr(1, lsData, chrCR)
    Loop
    
    If lsMSGflag = "R" Then
        res = Proc_Result(lsOrder, vasResTemp)
        
    ElseIf lsMSGflag = "Q" Then
        res = Proc_Order(lsRequest)
    End If

End Sub

Function Proc_Order(asReq As String) As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim iStr As Integer
    Dim iCnt As String
    
    Dim OKFlag As Integer
    
    Dim lsData As String
    Dim lsTemp As String
    
    Dim lsSampleNo As String
    Dim lsID As String
    Dim lsSampleType As String
    Dim lsRackID As String
    Dim lsPosNO As String
    Dim lsKind As String
    Dim lsPriority As String
    
    Dim lsCurDate As String
    
    Dim iRow As Integer
    
    lsData = asReq
    
    lsCurDate = Format(Date, "yyyymmdd") & Format(Time, "hhnnss")
    OKFlag = -1
    Proc_Order = -1
    
    gOrd.OrderCnt = 0
    gOrd.OrderText = ""
    gOrd.ExamCode = ""
    
    i = 0
    iStr = 1
    iCnt = 0
    
    i = InStr(iStr, lsData, "|")
    Do While i > 0
        iCnt = iCnt + 1
        
        lsTemp = Mid(lsData, iStr, i - iStr)
        lsData = Mid(lsData, i + 1)
        
        If iCnt = 3 Then
            OKFlag = 1
            Exit Do
        End If
        lsTemp = ""
        i = InStr(iStr, lsData, "|")
    Loop
    If OKFlag = 1 Then
        lsData = lsTemp

        i = InStr(1, lsData, "/")
        If i > 2 Then
            lsSampleNo = Mid(lsData, 3, i - 3)
            lsData = Mid(lsData, i + 1)
        End If

        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsID = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If

        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsSampleType = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If

        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsRackID = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If

        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsPosNO = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If

        i = InStr(1, lsData, "/")
        If i > 0 Then
            lsKind = Mid(lsData, 1, i - 1)
            lsData = Mid(lsData, i + 1)
        End If
        lsPriority = Trim(lsData)
        
'New Mode
'    If OKFlag = 1 Then
'        lsData = lsTemp
'
'        i = InStr(1, lsData, "^")
'        If i > 2 Then
'            lsSampleNo = Mid(lsData, 3, i - 3)
'            lsData = Mid(lsData, i + 1)
'        End If
'
'        i = InStr(1, lsData, "^")
'        If i > 0 Then
'            lsID = Mid(lsData, 1, i - 1)
'            lsData = Mid(lsData, i + 1)
'        End If
'
'        i = InStr(1, lsData, "^")
'        If i > 0 Then
'            lsSampleType = Mid(lsData, 1, i - 1)
'            lsData = Mid(lsData, i + 1)
'        End If
'
'        i = InStr(1, lsData, "^")
'        If i > 0 Then
'            lsRackID = Mid(lsData, 1, i - 1)
'            lsData = Mid(lsData, i + 1)
'        End If
'
'        i = InStr(1, lsData, "^")
'        If i > 0 Then
'            lsPosNO = Mid(lsData, 1, i - 1)
'            lsData = Mid(lsData, i + 1)
'        End If
'
'        i = InStr(1, lsData, "^")
'        If i > 0 Then
'            lsKind = Mid(lsData, 1, i - 1)
'            lsData = Mid(lsData, i + 1)
'        End If
'        lsPriority = Trim(lsData)

        iRow = -1
        For j = vasID.DataRowCnt To 1 Step -1
            If Trim(GetText(vasID, j, colBarCode)) = Trim(lsID) Then
                iRow = j
                Exit For
            End If
        Next j
        
        If iRow = -1 Then
            iRow = vasID.DataRowCnt + 1
            If iRow > vasID.MaxRows Then
                vasID.MaxRows = iRow + 1
            End If
        End If
        
        vasID.SetText colBarCode, iRow, Trim(lsID)
        vasID.SetText colRack, iRow, Trim(lsRackID)
        vasID.SetText colPos, iRow, Trim(lsPosNO)
'        vasID.SetText colSampleNo, iRow, lsSampleNo
'        vasID.SetText colSampleType, iRow, lsSampleType
        'vasID.SetText colKind, iRow, lsKind
        'vasID.SetText colPriority, iRow, lsPriority
        vasID.SetText colOrd, iRow, "0"
        
        gOrd.SampleType1 = lsSampleType
        res = MakeOrderRecode(lsID, lsPriority, Trim(lsRackID) & "-" & Trim(lsPosNO), lsKind, iRow)
        If res > 0 Then
            vasID.SetText colOrd, iRow, gOrd.OrderCnt
            vasID.SetText colState, iRow, "오더"
            Proc_Order = 1
        Else
            
            Proc_Order = 0
        End If
        
        If gOrd.SampleType1 = "0" Then
            Select Case gOrd.SampleType2
            Case "1", "2", "3", "4", "5"
                lsSampleType = gOrd.SampleType2
            Case Else
                lsSampleType = "1"
            End Select
        End If
        
        gOrderCnt = 1
        'lsSampleType = "1"
        If gOrd.OrderCnt > 0 Then
            gOrderMessage = "H|\^&|||host^2|||||H7600|TSDWN^BATCH|P|1" & chrCR & _
                            "P|1" & chrCR & _
                            "O|1|" & lsSampleNo & "^" & SetSpace(lsID, 22, 1) & "^" & lsSampleType & "^" & lsRackID & "^" & lsPosNO & "|" & lsKind & "|" & gOrd.OrderText & "|" & lsPriority & "||" & lsCurDate & "||||N||^^||||||^^^^||||||O" & chrCR & _
                            "L|1|N" & chrCR
                            '& chrETX
            'gOrderMessage = chrSTX & gOrderMessage & CheckSum(gOrderMessage) & chrCR & chrLF
    
            comState = "stTX"
        Else
            gOrderMessage = "H|\^&|||host^2|||||H7600|TSDWN^BATCH|P|1" & chrCR & _
                            "P|1" & chrCR & _
                            "O|1|" & lsSampleNo & "^" & SetSpace(lsID, 22, 1) & "^" & lsSampleType & "^" & lsRackID & "^" & lsPosNO & "|" & lsKind & "|^^^ALL|" & lsPriority & "||" & lsCurDate & "||||N||^^||||||^^^^||||||O" & chrCR & _
                            "L|1|N" & chrCR
                            '& chrETX
            'gOrderMessage = chrSTX & gOrderMessage & CheckSum(gOrderMessage) & chrCR & chrLF
            
            SaveQuery gOrderMessage, 1
            comState = "stTX"

        End If
        
        vasActiveCell vasID, iRow, colBarCode
    Else
        Proc_Order = 0
    End If
End Function

Public Function MakeOrderRecode(argCode As String, asEM As String, asRackPos As String, asKind As String, ByVal asRow As Long) As Integer
    Dim i, j As Integer
    Dim iCnt As Integer
    
    Dim retOrder As String
    Dim lsID As String
    Dim lsEquipCode As String
    Dim lsExamCode As String
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsSampleType As String
    Dim lsGubun As String       '진료,검진 여부
    Dim iISE As Integer
    Dim iISE_r As String
    
    Dim eDate As String
    
    Dim sCnt As String

    ClearSpread vasRes
    
    iCnt = 0
    MakeOrderRecode = -1
    
    gOrd.OrderCnt = 0
    gOrd.OrderText = ""
    gOrd.ExamCode = ""
    gOrd.SampleType2 = ""
    
    retOrder = ""
    ClearSpread vasTemp
    
    If argCode = "" Then
        MakeOrderRecode = -1
        Exit Function
    End If
    
    eDate = Trim(Text_Today.Text)

    lsID = Trim(argCode)
    
    '환자정보 불러오기
    'If Trim(GetText(vasID, asRow, colPName)) = "" Then
        Get_Sample_Info asRow
    'End If
    
    '처음 검사 샘플

    iISE = -1
    If vasOrder.DataRowCnt > 0 Then
        retOrder = ""

        For i = 1 To vasOrder.DataRowCnt
            lsEquipCode = ""
            lsExamName = ""
            lsSeqNo = "0"

            lsExamCode = Trim(GetText(vasOrder, i, 2))
            
            lsEquipCode = ""
            lsEquipCode = GetEquip_ExamCode(Trim(lsExamCode))
            If Trim(lsEquipCode) <> "" Then
                'retOrder = retOrder & "^^^" & lsEquipCode & "/\"
        
                If Trim(retOrder) = "" Then
                    retOrder = "^^^" & lsEquipCode & "/"
                Else
                    retOrder = retOrder & "\^^^" & lsEquipCode & "/"
                End If

                iCnt = iCnt + 1
            End If

        Next i

        If iISE = 1 Then
            iCnt = iCnt + 1
            If vasRes.MaxRows < iCnt Then vasRes.MaxRows = iCnt
            
            vasRes.SetText 1, iCnt, lsID
            vasRes.SetText colEquipCode, iCnt, "989"
            vasRes.SetText colExamCode, iCnt, ""
            vasRes.SetText colExamName, iCnt, "Na"
            'vasRes.SetText colSeqNo, iCnt, "0"

            Save_Local_One_1 asRow, iCnt, "A"

            iCnt = iCnt + 1
            If vasRes.MaxRows < iCnt Then vasRes.MaxRows = iCnt
            
            vasRes.SetText 1, iCnt, lsID
            vasRes.SetText colEquipCode, iCnt, "990"
            vasRes.SetText colExamCode, iCnt, ""
            vasRes.SetText colExamName, iCnt, "K"
            'vasRes.SetText colSeqNo, iCnt, "0"

            Save_Local_One_1 asRow, iCnt, "A"

            iCnt = iCnt + 1
            If vasRes.MaxRows < iCnt Then vasRes.MaxRows = iCnt
            vasRes.SetText 1, iCnt, lsID
            vasRes.SetText colEquipCode, iCnt, "991"
            vasRes.SetText colExamCode, iCnt, ""
            vasRes.SetText colExamName, iCnt, "Cl"
            'vasRes.SetText colSeqNo, iCnt, "0"

            Save_Local_One_1 asRow, iCnt, "A"


            If Trim(retOrder) = "" Then
                retOrder = "^^^989/\^^^990/\^^^991/"
            Else
                retOrder = retOrder & "\^^^989/\^^^990/\^^^991/"
            End If
        End If
    Else
        MakeOrderRecode = 0
    End If

    gOrd.OrderText = retOrder
    SaveQuery retOrder, 1
    
    vasTemp.SetText 2, 1, gOrd.OrderText
    
    gOrd.OrderCnt = iCnt
    gOrd.ExamCode = lsExamCode
    
    MakeOrderRecode = 1

End Function

Function Proc_Result(asOrd As String, ByVal argSpread As vaSpread) As Integer
    Dim i, j, k, iArr, lResRow As Long
    Dim iStr As Integer
    Dim iCnt As Integer
    Dim liRet As Integer
    
    Dim lsData As String
    Dim lsTemp As String
    Dim lsSampleType As String
    Dim lsSpecimenID As String
    Dim lsSpecimenID1 As String
    Dim lsOrder As String
    Dim lsID As String
    Dim lsRackID As String
    Dim lsPosNO As String
    Dim lsPriority As String
    
    Dim lsExamCode As String
    Dim lsExamDate As String
    Dim lsEquipCode As String
    Dim lsResult As String
    Dim lsEquipRes As String
    
    Dim lsUnit As String
    Dim lsRef As String
    Dim lsState As String
    Dim lsComment As String
    Dim lsPoint As String
    Dim sTmpStr As String
    
    Dim iRow As Integer
    
    Dim sCnt As String
    
    Dim lsExamName As String
    Dim lsSeqNo As String
    Dim lsGubun As String
    
    Dim sExamCodeAll As String
    
    Proc_Result = -1
    
'    gOrd.OrderCnt = 0
'    gOrd.OrderText = ""
    lsData = asOrd
    i = 0
    iStr = 1
    iCnt = 0
    lsID = ""
    i = InStr(iStr, lsData, "|")
    Do While i > 0
        iCnt = iCnt + 1
        
        lsTemp = Mid(lsData, iStr, i - iStr)
        lsData = Mid(lsData, i + 1)
        
        Select Case iCnt
        Case 3
            lsSpecimenID = lsTemp
        Case 4
            lsSpecimenID1 = lsTemp
        Case 5
            lsOrder = lsTemp
        Case 6
            lsPriority = lsTemp
            'Exit Do
        Case 23
            lsExamDate = lsTemp
            Exit Do
        Case Else
        End Select
        
        lsTemp = ""
        i = InStr(iStr, lsData, "|")
    Loop
    
'    'lsExamDate = Left(lsExamDate, 4) & "-" & Mid(lsExamDate, 5, 2) & "-" & Mid(lsExamDate, 7, 2) & " " & Mid(lsExamDate, 9, 2) & ":" & Mid(lsExamDate, 11, 2) & ":" & Mid(lsExamDate, 13, 2)
'    lsExamDate = Format(CDate(GetDateFull), "yyyy-mm-dd hh:nn:ss")
    
    i = InStr(1, lsSpecimenID, "^")
    If i > 0 Then
        lsSpecimenID = Mid(lsSpecimenID, i + 1)
        i = InStr(1, lsSpecimenID, "^")
        If i > 0 Then
            lsID = Trim(Left(lsSpecimenID, i - 1))
            lsSpecimenID = Mid(lsSpecimenID, i + 1)
            i = InStr(1, lsSpecimenID, "^")
            If i > 0 Then
                lsSampleType = Trim(Left(lsSpecimenID, i - 1))
                lsSpecimenID = Mid(lsSpecimenID, i + 1)
                i = InStr(1, lsSpecimenID, "^")
                If i > 0 Then
                    lsRackID = Left(lsSpecimenID, i - 1)
                    lsPosNO = Trim(Mid(lsSpecimenID, i + 1))
                End If
            End If
        End If
    End If
    
    'QC인 경우
    i = InStr(1, lsSpecimenID1, "^")
    If i > 0 Then
        'lsID = Trim(Left(lsSpecimenID1, i - 1))
        lsID = Trim(lsSpecimenID)
        lsSpecimenID1 = Mid(lsSpecimenID1, i + 1)
        i = InStr(1, lsSpecimenID1, "^")
        If i > 0 Then
            lsRackID = Left(lsSpecimenID1, i - 1)
            lsSpecimenID1 = Mid(lsSpecimenID1, i + 1)
            i = InStr(1, lsSpecimenID1, "^")
            If i > 0 Then
                lsPosNO = Trim(Left(lsSpecimenID1, i - 1))
                lsSpecimenID1 = Mid(lsSpecimenID1, i + 1)
                i = InStr(1, lsSpecimenID1, "^")
                If i > 0 Then
                    lsSpecimenID1 = Mid(lsSpecimenID1, i + 1)
                End If
                
                i = InStr(1, lsSpecimenID1, "^")
                If i > 0 Then
                    lsGubun = Trim(Left(lsSpecimenID1, i - 1))
                End If
            End If
        End If
    End If
    
    lsID = Trim(Mid(lsID, 1, 15))
    
    iRow = -1
    For i = vasID.DataRowCnt To 1 Step -1
        If Trim(GetText(vasID, i, colBarCode)) = lsID Then
            iRow = i
            Exit For
        End If
    Next i
    If iRow = -1 Then
        
        iRow = vasID.DataRowCnt + 1
        If iRow > vasID.MaxRows Then
            vasID.MaxRows = iRow + 1
        End If

    End If
    
    vasActiveCell vasID, iRow, colPID
         
    ClearSpread vasRes, 1, 1
    sExamCodeAll = ""
    
    'SetText vasID, lsPID, llRow, colPID
    SetText vasID, lsID, iRow, colBarCode
    
    vasID.SetText colRack, iRow, lsRackID
    vasID.SetText colPos, iRow, lsPosNO
    
    'vasID.SetText colExamDate, iRow, lsExamDate
    
    '환자정보 가져오기
    If Trim(GetText(vasID, iRow, colPName)) = "" Then
        Get_Sample_Info iRow
    End If
    
    '수신중========================================================
    SetText vasID, "수신중", iRow, colState
    SetBackColor vasID, iRow, iRow, 1, 1, 255, 250, 205
    '==============================================================
    
    For iArr = 1 To argSpread.DataRowCnt
        iStr = 1
        iCnt = 0
        lsData = GetText(argSpread, iArr, 1)
        If lsData <> "" Then
            i = InStr(iStr, lsData, "|")
            Do While i > 0
                iCnt = iCnt + 1
                lsTemp = Mid(lsData, iStr, i - iStr)
                lsData = Mid(lsData, i + 1)
                
                Select Case iCnt
                Case 3
                    lsEquipCode = lsTemp
                    j = InStr(1, lsEquipCode, "/")
                    If j > 0 Then
                        lsEquipCode = Mid(lsEquipCode, 4, j - 4)
                    Else
                    lsEquipCode = Mid(lsEquipCode, 4, Len(lsEquipCode) - 4)
                    End If
                Case 4
                    lsResult = lsTemp
                    lsEquipRes = lsResult
                    If InStr(1, lsResult, "^") > 0 Then
                        lsResult = Mid(lsResult, InStr(1, lsResult, "^") + 1)
                        lsEquipRes = lsResult
                    End If
                    'Exit Do
                Case 5
                    lsUnit = lsTemp
                Case 7
                    lsRef = lsTemp
                    If UCase(lsRef) = "N" Then lsRef = ""
                    Exit Do
                Case 9
                    lsState = lsTemp
                    Exit Do
                Case Else
                End Select
                
                lsTemp = ""
                i = InStr(iStr, lsData, "|")
            
            Loop
            
            lResRow = iArr
                        
            lsExamCode = ""
            
            If vasRes.MaxRows < lResRow Then vasRes.MaxRows = lResRow
            
            vasRes.SetText colEquipCode, lResRow, lsEquipCode
            vasRes.SetText colResult1, lResRow, lsEquipRes
            vasRes.SetText colResult, lResRow, lsResult

            SQL = "Select examcode, examname, resprec, seqno From equipexam" & vbCrLf & _
                  " Where Equipno = '" & gEquip & "' " & vbCrLf & _
                  "  And equipcode = '" & lsEquipCode & "'" '& vbCrLf & _
                  "  and examcode in (" & sExamCodeAll & ") "
            res = db_select_Col(gLocal, SQL)
            If (res = 1) And (gReadBuf(0) <> "") Then
                'j = vasRes.DataRowCnt + 1
        
'                If IsNumeric(lsResult) Then
                    lsExamCode = Trim(gReadBuf(0))
                    lsExamName = Trim(gReadBuf(1))
                    lsPoint = Trim(gReadBuf(2))
            
                    '2009.11.12 이상은
                    '소수점 처리
                    If IsNumeric(lsPoint) = True And IsNumeric(lsResult) = True Then
                        If CInt(lsPoint) > 0 Then
                            sTmpStr = "#0."
                            For i = 1 To CInt(lsPoint)
                                sTmpStr = sTmpStr & "0"
                            Next i
                        ElseIf CInt(lsPoint) = 0 Then
                            sTmpStr = "#0"
                        Else
                            sTmpStr = ""
                        End If
                        If Trim(sTmpStr) <> "" Then
                            lsResult = Format(lsResult, sTmpStr)
                        End If
                    End If

                '***********************************************
                res = Result_Set(lsEquipCode, lsResult)
                If res > 0 Then
                    lResRow = res
                    Save_Local_One_1 iRow, lResRow, "A"
                End If
                '***********************************************
            Else
                Save_Local_One_1 iRow, lResRow, "A"
            End If
        
                
        End If
    Next iArr
    
    
    If chkMode.Value = 1 Then
        liRet = -1
        
        liRet = Insert_Data(iRow)
        If liRet = 1 Then
            SetBackColor vasID, iRow, iRow, colCheckBox, colState, 202, 255, 112
            SetText vasID, "전송", iRow, colState
        ElseIf liRet = -1 Then
            SetForeColor vasID, iRow, iRow, colCheckBox, colState, 255, 0, 0
            SetText vasID, "실패", iRow, colState
        End If
    Else
        '수신중========================================================
        SetText vasID, "결과", iRow, colState
        SetBackColor vasID, iRow, iRow, 1, 1, 0, 128, 64
        '==============================================================
    End If
    
End Function

Function Result_Set(ByVal asTest As String, ByVal asRes As String) As Integer
    Dim sGiho As String
    Dim sRes As String
    Dim sRes1 As String
    Dim sFormat As String
    Dim sExamCode As String
    Dim sExamName As String
    Dim sValFlag As String
    
    Dim iRCnt
    
    Dim i As Integer
    Dim lResRow As Integer
    
    Result_Set = -1
    
    If Trim(asTest) = "" Then Exit Function
    
    SQL = "Select EquipCode, ExamCode, ExamName, resgubun, resprec, CutOffFlag, " & vbCrLf & _
          " NegValue, NegEqual, PosValue, PosEqual, cutoff" & vbCrLf & _
          "from EquipExam " & vbCrLf & _
          "where Equipno = '" & gEquip & "' " & vbCrLf & _
          "  and EquipCode = '" & Trim(asTest) & "' "
    res = db_select_Col(gLocal, SQL)
    If res < 1 Then Exit Function
    If Trim(gReadBuf(0)) <> Trim(asTest) Then Exit Function
    
    sGiho = ""
    sRes = ""
    sRes1 = ""
    
    sExamCode = Trim(gReadBuf(1))
    sExamName = Trim(gReadBuf(2))
    sValFlag = Trim(gReadBuf(10))
    
    If Trim(sExamCode) = "" Then Exit Function
    
    For i = 1 To Len(asRes)
        If IsNumeric(Mid(asRes, i, 1)) = True Or Mid(asRes, i, 1) = "." Then
            sRes = sRes & Mid(asRes, i, 1)
        Else
            sGiho = sGiho & Mid(asRes, i, 1)
        End If
    Next i
    
    Select Case Trim(gReadBuf(3))
    Case "I"
        sRes1 = Format(CCur(sRes), "#0")
        sRes1 = sGiho & sRes1
    Case "F"
        sFormat = ""
        For i = 1 To CInt(gReadBuf(4))
            sFormat = sFormat & "0"
        Next i
        sFormat = "0." & sFormat
        sRes1 = Format(CCur(sRes), sFormat)
        
        sRes1 = sGiho & sRes1
    Case "T"
'        sRes = ""
'
'        For i = 1 To Len(sResult)
'            If IsNumeric(Mid(sResult, i, 1)) = True Or Mid(sResult, i, 1) = "." Then
'                sRes = sRes & Mid(sResult, i, 1)
'            Else
'                sGiho = sGiho & Mid(sResult, i, 1)
'            End If
'        Next i
'
'        sFormat = ""
'        For i = 1 To CInt(gReadBuf(4))
'            sFormat = sFormat & "0"
'        Next i
'        sFormat = "0." & sFormat
'        sRes1 = Format(CCur(sRes), sFormat)
        
'        sRes1 = sGiho & sRes
'
'        sRes1 = UCase(sResultT) & "(" & sRes1 & ")"
        
'        'CuttOff
        If Trim(gReadBuf(5)) = "1" Then     '크다
            If Trim(gReadBuf(7)) = "1" And Trim(gReadBuf(9)) = "1" Then
                If CCur(sRes) <= CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) >= CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            ElseIf Trim(gReadBuf(7)) = "1" And Trim(gReadBuf(9)) = "0" Then
                If CCur(sRes) <= CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) > CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            ElseIf Trim(gReadBuf(7)) = "0" And Trim(gReadBuf(9)) = "1" Then
                If CCur(sRes) < CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) >= CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            ElseIf Trim(gReadBuf(7)) = "0" And Trim(gReadBuf(9)) = "0" Then
                If CCur(sRes) < CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) > CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            End If
        ElseIf Trim(gReadBuf(5)) = "2" Then      '작다
            If Trim(gReadBuf(7)) = "1" And Trim(gReadBuf(9)) = "1" Then
                If CCur(sRes) >= CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) <= CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
                
            ElseIf Trim(gReadBuf(7)) = "1" And Trim(gReadBuf(9)) = "0" Then
                If CCur(sRes) >= CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) < CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
                
            ElseIf Trim(gReadBuf(7)) = "0" And Trim(gReadBuf(9)) = "1" Then
                If CCur(sRes) > CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) <= CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            ElseIf Trim(gReadBuf(7)) = "0" And Trim(gReadBuf(9)) = "0" Then
                If CCur(sRes) > CCur(Trim(gReadBuf(6))) Then
                    sRes1 = "NEG"
                ElseIf CCur(sRes) < CCur(Trim(gReadBuf(8))) Then
                    sRes1 = "POS"
                Else
                    sRes1 = "Weak-POS"
                End If
                If Trim(sValFlag) = "1" Then
                    sRes1 = sRes1 & "(" & sRes & ")"
                End If
                
            End If
        End If
    
    End Select
    
    lResRow = -1
    For i = 1 To vasRes.DataRowCnt
        If Trim(asTest) = Trim(GetText(vasRes, i, colEquipCode)) Then
            lResRow = i
            Exit For
        End If
    Next i
    
    If lResRow = -1 Then
        lResRow = vasRes.DataRowCnt + 1
        If lResRow > vasRes.MaxRows Then
            vasRes.MaxRows = lResRow
        End If
    End If
    
    SetText vasRes, gsBarCode, lResRow, colBarCode      '검체번호
    SetText vasRes, asTest, lResRow, colEquipCode       '장비코드
    SetText vasRes, sExamCode, lResRow, colExamCode     '검사코드
    SetText vasRes, sExamName, lResRow, colExamName     '검사명
    SetText vasRes, sRes1, lResRow, colResult           '검사결과
    SetText vasRes, asRes, lResRow, colResult1          '장비결과
    
'    If IsNumeric(GetText(vasID, glRow, colRCnt)) Then
'        iRCnt = CInt(GetText(vasID, glRow, colRCnt)) + 1
'        SetText vasID, CStr(iRCnt), glRow, colRCnt
'    Else
'        SetText vasID, "1", glRow, colRCnt
'    End If

    SetText vasID, vasRes.DataRowCnt, glRow, colRes
    If InStr(1, Trim(GetText(vasRes, lResRow, colResult)), "POS") > 0 Then
        vasRes.Row = lResRow
        vasRes.Col = colResult
        vasRes.ForeColor = RGB(205, 55, 0)
    ElseIf InStr(1, Trim(GetText(vasRes, lResRow, colResult)), "Weak-POS") > 0 Then
        vasRes.Row = lResRow
        vasRes.Col = colResult
        vasRes.ForeColor = RGB(55, 0, 205)
    Else
        vasRes.Row = lResRow
        vasRes.Col = colResult
        vasRes.ForeColor = RGB(0, 0, 0)
    End If
    
    Result_Set = lResRow
End Function

Function Save_Local_One_1(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String

    sExamDate = GetDateFull

    sCnt = ""
    If Trim(GetText(vasRes, asRow2, colEquipCode)) = "" Then Exit Function
    
    SQL = "select count(*) from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "'" & vbCrLf & _
          "  and examcode= '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "'"
    res = db_select_Col(gLocal, SQL)
    sCnt = Trim(gReadBuf(0))
    If res = -1 Then
        SaveQuery SQL, 1
        Exit Function
    End If
    
    If Not IsNumeric(sCnt) Then
        sCnt = "0"
    End If
    
    If Not IsNumeric(GetText(vasID, asRow1, colPAge)) Then
        SetText vasID, "0", asRow1, colPAge
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
'        SetText vasExam, "1900-01-01", asRow, colExamDate
'    End If

    If sCnt = "0" Then
        SQL = "INSERT INTO pat_res (examdate, equipno, barcode, seqno, diskno, posno, " & _
              "pid, pname, jumin, page, psex, resdate, receno, " & _
              "equipcode, examcode, result, result1, sendflag, examname, " & _
              "refflag, refvalue, panicvalue, recedate ) " & vbCrLf & _
              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
              "'" & Trim(GetText(vasID, asRow1, colBarCode)) & "', '" & Trim(GetText(vasID, asRow1, colSeqNo)) & "'," & _
              "'" & Trim(GetText(vasID, asRow1, colRack)) & "', '" & Trim(GetText(vasID, asRow1, colPos)) & "', " & _
              "'" & Trim(GetText(vasID, asRow1, colPID)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasID, asRow1, colPName)) & "', '" & Trim(GetText(vasID, asRow1, colPJumin)) & "', " & _
              "'" & Trim(GetText(vasID, asRow1, colPAge)) & "', '" & Trim(GetText(vasID, asRow1, colPSex)) & "', " & _
              "'" & sExamDate & "', '" & Trim(GetText(vasID, asRow1, colReceno)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "',  " & _
              "'" & Trim(GetText(vasRes, asRow2, colResult)) & "', '" & Trim(GetText(vasRes, asRow2, colResult1)) & "', '" & asSend & "', '" & Trim(GetText(vasRes, asRow2, colExamName)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasRes, asRow2, colRCheck)) & "',  " & _
              "'" & Trim(GetText(vasID, asRow1, colOrd)) & "', '" & Trim(GetText(vasID, asRow1, colRes)) & "', '" & Trim(GetText(vasID, asRow1, colDate)) & "') "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    Else
        SQL = " Update pat_res Set " & vbCrLf & _
              " diskno = '" & Trim(GetText(vasID, asRow1, colRack)) & "', " & vbCrLf & _
              " posno  = '" & Trim(GetText(vasID, asRow1, colPos)) & "', " & vbCrLf & _
              " result = '" & Trim(GetText(vasRes, asRow2, colResult)) & "', " & vbCrLf & _
              " result1 = '" & Trim(GetText(vasRes, asRow2, colResult1)) & "', " & vbCrLf & _
              " refflag = '" & Trim(GetText(vasRes, asRow2, colRCheck)) & "', " & vbCrLf & _
              " refvalue = '" & Trim(GetText(vasID, asRow1, colOrd)) & "', " & vbCrLf & _
              " panicvalue = '" & Trim(GetText(vasID, asRow1, colRes)) & "', " & vbCrLf & _
              " resdate = '" & sExamDate & "' " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasID, asRow1, colBarCode)) & "' " & vbCrLf & _
              " And equipcode = '" & Trim(GetText(vasRes, asRow2, colEquipCode)) & "' " & vbCrLf & _
              " And examcode = '" & Trim(GetText(vasRes, asRow2, colExamCode)) & "' "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If

    End If
    
End Function

Function Save_Local_One_2(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String

    sExamDate = GetDateFull
    
    sCnt = ""
    'If Trim(GetText(vasOrderTmp, asRow2, colEquipCode)) = "" Then Exit Function
    
    SQL = "select count(*) from pat_res " & vbCrLf & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & Trim(GetText(vasIDTmp, asRow1, colBarCode)) & "' " & vbCrLf & _
          "  and equipcode = '" & Trim(GetText(vasOrderTmp, asRow2, colEquipCode)) & "'" & vbCrLf & _
          "  and examcode= '" & Trim(GetText(vasOrderTmp, asRow2, colExamCode)) & "'"
    res = db_select_Col(gLocal, SQL)
    sCnt = Trim(gReadBuf(0))
    If res = -1 Then
        SaveQuery SQL, 1
        Exit Function
    End If
    
    If Not IsNumeric(sCnt) Then
        sCnt = "0"
    End If
    
    If Not IsNumeric(GetText(vasIDTmp, asRow1, colPAge)) Then
        SetText vasIDTmp, "0", asRow1, colPAge
    End If
'    If Not IsDate(Trim(GetText(vasExam, asRow, colExamDate))) Then
'        SetText vasExam, "1900-01-01", asRow, colExamDate
'    End If

    If sCnt = "0" Then
        SQL = "INSERT INTO pat_res (examdate, equipno, barcode, seqno, diskno, posno, " & _
              "pid, pname, jumin, page, psex, resdate, receno, " & _
              "equipcode, examcode, result, result1, sendflag, examname, " & _
              "refflag, refvalue, panicvalue, recedate ) " & vbCrLf & _
              "VALUES ('" & Format(CDate(Text_Today.Text), "yyyymmdd") & "', '" & Trim(gEquip) & "', " & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colBarCode)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colSeqNo)) & "'," & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colRack)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colPos)) & "', " & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colPID)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colPName)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colPJumin)) & "', " & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colPAge)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colPSex)) & "', " & _
              "'" & sExamDate & "', '" & Trim(GetText(vasIDTmp, asRow1, colReceno)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasOrderTmp, asRow2, colEquipCode)) & "', '" & Trim(GetText(vasOrderTmp, asRow2, colExamCode)) & "',  " & _
              "'" & Trim(GetText(vasOrderTmp, asRow2, colResult)) & "', '" & Trim(GetText(vasOrderTmp, asRow2, colResult1)) & "', '" & asSend & "', '" & Trim(GetText(vasOrderTmp, asRow2, colExamName)) & "', " & vbCrLf & _
              "'" & Trim(GetText(vasOrderTmp, asRow2, colRCheck)) & "',  " & _
              "'" & Trim(GetText(vasIDTmp, asRow1, colOrd)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colRes)) & "', '" & Trim(GetText(vasIDTmp, asRow1, colDate)) & "') "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If
    Else
        SQL = " Update pat_res Set " & vbCrLf & _
              " diskno = '" & Trim(GetText(vasIDTmp, asRow1, colRack)) & "', " & vbCrLf & _
              " posno  = '" & Trim(GetText(vasIDTmp, asRow1, colPos)) & "', " & vbCrLf & _
              " result = '" & Trim(GetText(vasOrderTmp, asRow2, colResult)) & "', " & vbCrLf & _
              " result1 = '" & Trim(GetText(vasOrderTmp, asRow2, colResult1)) & "', " & vbCrLf & _
              " refflag = '" & Trim(GetText(vasOrderTmp, asRow2, colRCheck)) & "', " & vbCrLf & _
              " refvalue = '" & Trim(GetText(vasIDTmp, asRow1, colOrd)) & "', " & vbCrLf & _
              " panicvalue = '" & Trim(GetText(vasIDTmp, asRow1, colRes)) & "', " & vbCrLf & _
              " resdate = '" & sExamDate & "' " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasIDTmp, asRow1, colBarCode)) & "' " & vbCrLf & _
              " And equipcode = '" & Trim(GetText(vasOrderTmp, asRow2, colEquipCode)) & "' " & vbCrLf & _
              " And examcode = '" & Trim(GetText(vasOrderTmp, asRow2, colExamCode)) & "' "
        
        res = SendQuery(gLocal, SQL)
        If res = -1 Then
            SaveQuery SQL
            Exit Function
        End If

    End If
    
End Function

Function Insert_Data(ByVal argSpcRow As Long, Optional asSend As Integer = 0) As Integer
'서버의 데이타 베이스에 저장
    Dim sDpcd, sDate1, sSlip, sItem, sOitp, sWkno As String
    Dim sIDNo, sSmyr, sSmsn, sSms1 As String
    Dim tSmsn As String
    Dim lsExamCode, lsResult As String
    Dim lPanicLow, lPanicHigh As Currency
    Dim lDeltaLow, lDeltaHigh, lDeltaMeth, lDeltaGap
    Dim lsPanic, lsDelta As String
    Dim lsPreDate, lsPreResult As String
    Dim lsNState, lsWState As String
    Dim lStdVal
    Dim lTerm As Long
    Dim lsQCChk As String

    Dim iNone, iDP

    Dim sResDate As String
    Dim sRDate As String
    Dim sRTime As String

    Dim lsID As String

    Dim i, j As Long
    Dim lRow As Long
    Dim lsQCOn As String
    
    Dim sResult As String
    Dim sExamCode As String
    Dim sBarCode As String
    Dim sEquipCode As String
    Dim sResStr As String
    Dim sResRow As Long
    Dim sResCnt As String
    Dim sEquipRes As String
    Dim sParam As String
    Dim X As Integer
    
    Dim sReviewRes As String
    
    Insert_Data = -1

    lsQCOn = ""

    lRow = argSpcRow

    If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Function

    lsID = Trim(GetText(vasID, lRow, colBarCode))
    sBarCode = ""
    sEquipCode = ""
    sResult = ""
    sExamCode = ""
    
    If lsID = "" Then Exit Function

    ClearSpread vasTemp
    ClearSpread vasTemp1

    iNone = 0
    iDP = 0

    SQL = "Select equipcode, examcode, examname, result, result1 " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' And examdate = '" & Format(Text_Today.Text, "YYYYMMDD") & "' " & vbCrLf & _
          "  and barcode = '" & lsID & "' " & vbCrLf & _
          "  and examcode <> '' " '& vbCrLf & _
          "  and result <> '' "
    If asSend = 0 Then
'        SQL = SQL & vbCrLf & _
'          "  and sendflag <> 'C' "
    End If
    
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If vasTemp.DataRowCnt < 1 Then Exit Function

    Save_Raw_Data lsID & " : 서버 결과 전송 시작"

    On Error GoTo ErrHandle
    
    sReviewRes = ""
    
    ClearSpread vasTemp2
    
    SQL = " Select b.examname, a.result From pat_res a, equipexam b " & vbCrLf & _
          "where a.equipno = '" & gEquip & "' And a.examdate = '" & Format(Text_Today.Text, "YYYYMMDD") & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and a.result = '+' and a.equipcode = b.equipcode and a.examcode = b.examcode "
    res = db_select_Vas(gLocal, SQL, vasTemp2)
    
    For i = 1 To vasTemp2.DataRowCnt
        If sReviewRes = "" Then
            sReviewRes = "REVIEW/" & Trim(GetText(vasTemp2, i, 1))
        Else
            sReviewRes = sReviewRes & "/" & Trim(GetText(vasTemp2, i, 1))
        End If
    Next i
    
    sParam = ""
    
    For sResRow = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, sResRow, 2)) <> "" Then
            sResult = ""
            If Trim(GetText(vasTemp, sResRow, 1)) = "A000" And Trim(GetText(vasTemp, sResRow, 5)) = "REV" Then
                sResult = sReviewRes
            Else
                sResult = Trim(GetText(vasTemp, sResRow, 4))
            End If
            
            Debug.Print Trim(GetText(vasTemp, sResRow, 2)) & " : " & sResult
            sParam = sParam & "<Table>" & _
                    "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                    "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                    "<USERID><![CDATA[LIA]]></USERID>" & _
                    "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                    "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                    "<P0><![CDATA[" & lsID & "]]></P0>" & _
                    "<P1><![CDATA[" & Trim(GetText(vasTemp, sResRow, 2)) & "]]></P1>" & _
                    "<P2><![CDATA[" & sResult & "]]></P2>" & _
                    "<P3><![CDATA[]]></P3>" & _
                    "<P4><![CDATA[" & gEquip & "]]></P4>" & _
                    "<P5><![CDATA[" & gIFUser & "]]></P5>" & _
                    "<P6><![CDATA[]]></P6>" & _
                    "<P7><![CDATA[]]></P7>" & _
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
    End If
    
    Insert_Data = 1

    Save_Raw_Data lsID & " : 서버 결과 전송 완료!"
    
    Exit Function

ErrHandle:
    Save_Raw_Data Err.Number & " : " & Err.Description & vbCrLf & _
                  SQL
    Resume Next
    
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

Private Sub Picture1_Click()
    frmUser.Show 0
End Sub

Private Sub Text_Today_GotFocus()
    SelectFocus Text_Today
End Sub

Private Sub Text_Today_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCall_Click
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
'        LX20 Text2

    End If
End Sub

Private Sub Timer1_Timer()
'    Dim sRet        As String
'    Dim sCnt        As String
'    Dim lRow        As Long
'    Dim lsRow       As Long
'
'    Dim i           As Integer
'    Dim j           As Integer
'    Dim lsPatFlag   As Boolean
'
'    Timer1.Enabled = False
'    Save_Raw_Data "[오더]"
'
'    '로컬에 오더 쌓아놓기
'    ClearSpread vasWork
'
'    'sRet = Online_TLA(gXml_S04, Text_Today.Text, Text_Today.Text)
'    sRet = Online_TLA(gXml_S04, "2010-01-09", Text_Today.Text)
'
'    For i = 0 To giIndex
'        lsRow = -1
'        lsPatFlag = False
'
'        If lsRow = -1 Then
'            lsRow = vasWork.DataRowCnt + 1
'        End If
'
'        If vasWork.MaxRows < lsRow Then
'            vasWork.MaxRows = lsRow
'        End If
'        If lsPatFlag = True Then
'
'        Else
'            vasWork.SetText 1, lsRow, gTLA_Info_Select(i).SPCNO
'            vasWork.SetText 2, lsRow, gTLA_Info_Select(i).TST_DTE
'        End If
'
'        vasWork.MaxRows = vasWork.DataRowCnt
'    Next i
'
'    For lRow = 1 To vasWork.DataRowCnt
'        If Trim(GetText(vasWork, lRow, 1)) <> "" Then
'            sCnt = "0"
'
'            SQL = " Select count(*) From pat_res " & CR & _
'                  " Where equipno = '" & gEquip & "' " & CR & _
'                  " And barcode = '" & Trim(GetText(vasWork, lRow, 1)) & "' " & CR & _
'                  " And sendflag = '0' "
'            res = db_select_Col(gLocal, SQL)
'
'            If gReadBuf(0) = "" Then
'                sCnt = "0"
'            Else
'                sCnt = gReadBuf(0)
'            End If
'
'            If sCnt = "0" Then
'                '환자정보 불러오기
'                ClearSpread vasIDTmp
'
'                sRet = Online_XML(gXml_S03, Trim(GetText(vasWork, lRow, 1)))
'
'                SetText vasIDTmp, Trim(GetText(vasWork, lRow, 1)), lRow, colBarCode
'                SetText vasIDTmp, gPat_Info_Select.PT_NO, lRow, colPID
'                SetText vasIDTmp, gPat_Info_Select.PT_NM, lRow, colPName
'                SetText vasIDTmp, gPat_Info_Select.SEX, lRow, colPSex
'                SetText vasIDTmp, gPat_Info_Select.AGE, lRow, colPAge
'                SetText vasIDTmp, gPat_Info_Select.ACPTNO_1, lRow, colSeqNo
'                'SetText vasIDTmp, gPat_Info_Select.ACPT_DTETM, lRow, colDate
'                SetText vasIDTmp, Mid(gPat_Info_Select.ACPT_DTETM, 1, 10), lRow, colDate
'                SetText vasIDTmp, gPat_Info_Select.SPC_CD_1, lRow, colReceno
'
'                ClearSpread vasOrderTmp
'
'                sRet = Online_XML(gXml_S07, Trim(GetText(vasWork, lRow, 1)))
'
'                For i = 0 To giIndex
'                    SetText vasOrderTmp, Trim(gExam_Select(i).TST_CD), i + 1, colExamCode
'
'                    SQL = " Select equipcode, examname From equipexam " & CR & _
'                          " Where equipno = '" & gEquip & "' " & CR & _
'                          " And examcode= '" & Trim(GetText(vasOrderTmp, i, colExamCode)) & "' "
'                    res = db_select_Col(gLocal, SQL)
'                    If res = 1 Then
'                        SetText vasOrderTmp, Trim(gReadBuf(0)), i + 1, colEquipCode
'                        SetText vasOrderTmp, Trim(gReadBuf(1)), i + 1, colExamName
'                    End If
'                Next i
'
'                For j = 1 To vasOrderTmp.DataRowCnt
'                    Save_Local_One_2 lRow, j, "0"
'                Next j
'            End If
'
'        End If
'    Next lRow
'
'    Save_Raw_Data "[타이머]"
'    Timer1.Enabled = True
End Sub

Private Sub txtEnd_GotFocus()
    SelectFocus txtEnd
End Sub

Private Sub txtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtEnd) = False Then
            txtEnd.SetFocus
            Exit Sub
        End If
        cmdSend.SetFocus
    End If
End Sub

Private Sub txtHelp_Change()

End Sub

Private Sub txtID_GotFocus()
    SelectFocus txtID
End Sub

Private Sub txtStart_GotFocus()
    SelectFocus txtStart
End Sub

Private Sub txtStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsNumeric(txtStart) = False Then
            txtStart.SetFocus
            Exit Sub
        End If
        txtEnd.SetFocus
    End If
End Sub


Private Sub vasID_Click(ByVal Col As Long, ByVal Row As Long)
    If Row = 0 Then
        If Col = colRack Or Col = colPos Then
            vasSort vasID, colRack, colPos
        Else
            vasSort vasID, Col
        End If
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
    Dim lsTime As String
    Dim lsState As String
    
    
    Dim iRow As Long
    
    'cmdCall_Click
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    
    lsID = Trim(GetText(vasID, Row, colBarCode))
    
'    If Trim(GetText(vasID, Row, colState)) = "결과" Then
'        lsState = "A"
'    ElseIf Trim(GetText(vasID, Row, colState)) = "완료" Then
'        lsState = "C"
'    End If
    'Local에서 불러오기
    ClearSpread vasRes
    
    If Trim(GetText(vasID, Row, colPJumin)) = "F" Then
        lsTime = Trim(GetText(vasID, Row, colPID))
        If Len(lsTime) = 4 Then
        Else
            lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
        End If
        SQL = "select a.equipcode, min(b.examcode), min(b.examname), a.result, b.seqno, a.resflag, a.result " & vbCrLf & _
              " From qc_res a, equipexam b " & vbCrLf & _
              "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
              "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
              "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
              "  and a.levelname = '" & lsID & "' " & vbCrLf & _
              "  and b.equipno = a.equipno " & vbCrLf & _
              "  and b.equipcode = a.equipcode " & vbCrLf & _
              "group by a.equipcode, a.result, b.seqno, a.resflag, a.result "
        res = db_select_Vas(gLocal, SQL, vasRes)
    End If
    

    '장비코드, 검사코드, 검사명, 결과, 순번
    SQL = "Select a.equipcode, a.examcode, b.examname, a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
          "from pat_res a, equipexam b " & vbCrLf & _
          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and a.examcode <> a.equipcode " & vbCrLf & _
          "  and b.equipno = a.equipno " & vbCrLf & _
          "  and b.equipcode = a.equipcode " & vbCrLf & _
          "  and b.examcode = a.examcode"
    res = db_select_Vas(gLocal, SQL, vasRes)
    SQL = "Select a.equipcode, a.examcode, max(b.examname), a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
          "from pat_res a, equipexam b " & vbCrLf & _
          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' " & vbCrLf & _
          "  and a.examcode = a.equipcode " & vbCrLf & _
          "  and b.equipno = a.equipno " & vbCrLf & _
          "  and b.equipcode = a.equipcode " & vbCrLf & _
          "group by a.equipcode, a.examcode, a.result, b.seqno, a.refflag, a.result1 "
    res = db_select_Vas(gLocal, SQL, vasRes, vasRes.DataRowCnt + 1, 1)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    For iRow = 1 To vasRes.DataRowCnt
        If Trim(GetText(vasRes, iRow, colRCheck)) <> "" Then
            SetForeColor vasRes, iRow, iRow, colResult, colResult, 255, 0, 0
        Else
            SetForeColor vasRes, iRow, iRow, colResult, colResult, 0, 0, 0
        End If
    Next iRow
    vasRes.MaxRows = vasRes.DataRowCnt
    'vasSort vasRes, 5, 2
End Sub

Private Sub vasID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Long
    Dim lsID As String
    Dim lsTime As String
    
    iRow = vasID.ActiveRow
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasID.DataRowCnt Then
            Exit Sub
        End If
        
        lsID = Trim(GetText(vasID, iRow, colBarCode))
        
        If Trim(GetText(vasID, iRow, colPJumin)) = "F" Then
            If MsgBox("해당 QC 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                Exit Sub
            End If
            
            lsTime = Trim(GetText(vasID, iRow, colPID))
            If Len(lsTime) = 4 Then
            Else
                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
            End If
            
            SQL = "Delete From qc_res a " & vbCrLf & _
                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
                  "  and a.levelname = '" & lsID & "' "
            res = SendQuery(gLocal, SQL)
                
            Exit Sub
        End If
            
        If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
            
        SQL = " Delete From pat_res " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
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

Private Sub vasID_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lRow As Long
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        lRow = vasID.ActiveRow
        If lRow < 1 Or lRow > vasID.DataRowCnt Then Exit Sub
            
        vasID_DblClick colBarCode, lRow
    End If
End Sub

Private Sub vasID_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'Dim iRow As Long
'Dim lsID As String
'
'    If MsgBox("해당 환자결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
'        Exit Sub
'    End If
'
'    iRow = Row
'
'    lsID = Trim(GetText(vasID, iRow, colBarcode))
'
'    SQL = " Delete From pat_res " & vbCrLf & _
'          " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
'          " And equipno = '" & gEquip & "' " & vbCrLf & _
'          " And barcode = '" & lsID & "' "
'    res = SendQuery(gLocal, SQL)
'
'    If res = -1 Then
'        SaveQuery SQL
'        Exit Sub
'    End If
'
'    DeleteRow vasID, iRow, iRow
End Sub

Private Sub vasRes_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vasResRow As Long
    Dim vasResCol As Long
    Dim vasIDRow As Long
        
    Dim lCCR, lM_C_ratio, lP_C_ratio As Long
    Dim sCCR, sCrea_S, sCrea_U, sM_ALB_U, sTP_U As String
    
    Dim sResult As String
    Dim sResult1 As String
    
    Dim i As Integer
    
    Dim sTotalVol As String
    
    Dim lsTime As String
    
    vasIDRow = vasID.ActiveRow
    vasResRow = vasRes.ActiveRow
    vasResCol = vasRes.ActiveCol
    
    If KeyCode = vbKeyReturn Then

        If vasResCol = colResult Then
            
            If Trim(GetText(vasRes, vasResRow, colEquipCode)) = "88888" Then
                sTotalVol = Trim(GetText(vasRes, vasResRow, colResult))
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                Save_Local_One_1 vasIDRow, vasResRow, "A"
            
            ElseIf Trim(GetText(vasRes, vasResRow, colEquipCode)) = "99999" Then
                sTotalVol = Trim(GetText(vasRes, vasResRow, colResult))
                SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
                Save_Local_One_1 vasIDRow, vasResRow, "A"
                
                If IsNumeric(sTotalVol) Then
                    lCCR = -1
                    sCCR = ""
                    sCrea_S = ""
                    sCrea_U = ""
                    sM_ALB_U = ""
                    sTP_U = ""
                    
                    i = 1
                    Do While i <= vasRes.DataRowCnt
                        Select Case Trim(GetText(vasRes, i, colExamCode))
                        Case "L3117", "L3101", "L3102", "L3103"  'Microalbumun(24hr),Na,K,Cl
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 1000, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                            
                        Case "L3104", "L3106", "L3107", "L3109" 'Ca,Pi,UA,Protein(24hr)
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31094", "L31095" 'Protein 16hr, 8hr
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31111", "L31112", "L31123", "L3113" 'Creatinie 16hr, 8hr,24hr, BUN(24hr UR)
                            sResult = Trim(GetText(vasRes, i, colResult1))
                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
                            'SetText vasRes, "L31123", i, colExamCode
                            'SetText vasRes, sResult, i, colResult1
                            If IsNumeric(sResult) Then
                                sResult = Format(CCur(sResult) * CCur(sTotalVol) / 100 / 1000, "0.00")
                                SetText vasRes, sResult, i, colResult
                            End If
                            
                            Save_Local_One_1 vasIDRow, i, "A"
                        Case "L3041", "88888"   'Serum Creatinine
                            sCrea_S = Trim(GetText(vasRes, i, colResult1))
                            
                            'Save_Local_One_1 vasIDRow, i, "A"
                        Case "L31121"   'CCR
                            sCCR = Trim(GetText(vasRes, i, colResult1))
                            lCCR = i
                        Case "L31171"   'Microalbumin(random)
                            sM_ALB_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31110"  'Creatinine(random)
                            sCrea_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31090"   'Protein(random)
                            sTP_U = Trim(GetText(vasRes, i, colResult1))
                        Case "L31172"   'Microalbumin / creatinine (random urine)
                            lM_C_ratio = i
                        Case "L31172"   'protein / creatinie (random)
                            lP_C_ratio = i
                        End Select
                        i = i + 1
                    Loop
                    
                    If lCCR > 0 And lCCR <= vasRes.DataRowCnt And IsNumeric(sCrea_U) = True And IsNumeric(sCrea_S) = True Then
                        sResult = Format(CCur(sCrea_U) * CCur(sTotalVol) / 1440 / CCur(sCrea_S), "0.000")
                        SetText vasRes, sResult, lCCR, colResult
                        SetText vasRes, sResult, lCCR, colResult1
                        Save_Local_One_1 vasIDRow, i, "A"
                    End If
                    
'                    If IsNumeric(sM_ALB_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sM_ALB_U) / CCur(sCrea_U), "0.00") * 100
'                        If lM_C_ratio > 0 And lM_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "101", i, colEquipCode
'                            SetText vasRes, "L31172", i, colExamCode
'                            SetText vasRes, "Microalbumin / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
'
'                    If IsNumeric(sTP_U) = True And IsNumeric(sCrea_U) = True Then
'                        sResult = Format(CCur(sTP_U) / CCur(sCrea_U), "0.00") * 1000
'                        If lP_C_ratio > 0 And lP_C_ratio <= vasRes.DataRowCnt Then
'                            SetText vasRes, sResult, lM_C_ratio, colResult
'                        Else
'                            i = vasRes.DataRowCnt + 1
'                            If i > vasRes.maxrows Then
'                                vasRes.maxrows = i
'                            End If
'
'                            SetText vasRes, "102", i, colEquipCode
'                            SetText vasRes, "L31201", i, colExamCode
'                            SetText vasRes, "Urine Protein / Urine Creatinine", i, colExamName
'                            SetText vasRes, sResult, i, colResult
'                            SetText vasRes, sResult, i, colResult1
'                        End If
'
'                        Save_Local_One_1 vasIDRow, i, "A"
'                    End If
                End If
            Else
                
                If Trim(GetText(vasRes, vasIDRow, colPJumin)) = "F" Then
                
                    If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 수정 하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                        Exit Sub
                    End If
                
                    lsTime = Trim(GetText(vasID, vasIDRow, colPID))
                    If Len(lsTime) = 4 Then
                    Else
                        lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
                    End If
                    
                    SQL = "update qc_res set result = '" & sResult & "' " & vbCrLf & _
                          "where equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                          "  and examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                          "  and examtime = '" & lsTime & "' " & vbCrLf & _
                          "  and levelname = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
                          "  and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
                    res = SendQuery(gLocal, SQL)
                
                    Exit Sub
                Else
                
                
                    sResult = Trim(GetText(vasRes, vasResRow, colResult))
                    If MsgBox("저장하시겠습니까?", vbYesNo + vbCritical + vbDefaultButton2, "주의!!!  확인!!!") = vbYes Then
                        sResult = Trim(GetText(vasRes, vasResRow, colResult))
                        
                        SQL = " update pat_res set " & vbCrLf & _
                              " Result = '" & sResult & "' " & vbCrLf & _
                              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
                              " And equipno = '" & gEquip & "' " & vbCrLf & _
                              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
                              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
                        res = SendQuery(gLocal, SQL)
                        
                        If res = -1 Then
                            SaveQuery SQL
                            Exit Sub
                        End If
        
                        'SetText vasRes, Trim(GetText(vasRes, vasResRow, colResult)), vasResRow, colResult1
        
                    End If
                End If
            End If
            
            
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If Trim(GetText(vasID, vasIDRow, colPJumin)) = "F" Then
        
            If MsgBox("해당 QC의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
                Exit Sub
            End If
        
            lsTime = Trim(GetText(vasID, vasIDRow, colPID))
            If Len(lsTime) = 4 Then
            Else
                lsTime = Left(lsTime, 2) & Mid(lsTime, 4, 2)
            End If
            
            SQL = "Delete From qc_res a " & vbCrLf & _
                  "where a.equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
                  "  and a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
                  "  and a.examtime = '" & lsTime & "' " & vbCrLf & _
                  "  and a.levelname = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
                  " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' "
            res = SendQuery(gLocal, SQL)
        
            Exit Sub
        End If
        If MsgBox("해당 환자의 " & Trim(GetText(vasRes, vasResRow, colExamName)) & " 결과를 삭제하시겠습니까?", vbInformation + vbYesNo, "알림") = vbNo Then
            Exit Sub
        End If
        
        SQL = " Delete From pat_res " & vbCrLf & _
              " Where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
              " And equipno = '" & gEquip & "' " & vbCrLf & _
              " And barcode = '" & Trim(GetText(vasID, vasIDRow, colBarCode)) & "' " & vbCrLf & _
              " and equipcode = '" & Trim(GetText(vasRes, vasResRow, colEquipCode)) & "' " & vbCrLf & _
              " and examcode =  '" & Trim(GetText(vasRes, vasResRow, colExamCode)) & "' "
        res = SendQuery(gLocal, SQL)
        
        If res = -1 Then
            SaveQuery SQL
            Exit Sub
        End If
            
        DeleteRow vasRes, vasResRow, vasResRow
    
    End If
End Sub

Function Save_Local_QC(asExamDate As String, asBarcode As String, asExamCode As String, asRes1 As String, asRes2 As String)
    Dim sResDateTime As String
    Dim sControl As String
    Dim sLotNo As String
    
    Dim sRefLow As String
    Dim sRefHigh As String
    Dim sRefFlag As String
    
    Dim sCnt As String
    
    sResDateTime = Format(CDate(asExamDate), "yyyymmdd hhnnss")
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



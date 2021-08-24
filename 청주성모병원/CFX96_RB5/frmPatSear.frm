VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmPatSear 
   BorderStyle     =   1  '단일 고정
   Caption         =   "환자조회"
   ClientHeight    =   8535
   ClientLeft      =   7440
   ClientTop       =   2250
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9450
   StartUpPosition =   2  '화면 가운데
   Begin VB.PictureBox sspOrder 
      BackColor       =   &H00FFC0C0&
      Height          =   4005
      Left            =   810
      ScaleHeight     =   3945
      ScaleWidth      =   7635
      TabIndex        =   6
      Top             =   2820
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   660
         TabIndex        =   31
         Top             =   3330
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   30
         Top             =   3330
         Width           =   375
      End
      Begin VB.TextBox txtNo 
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
         Height          =   345
         Left            =   1290
         TabIndex        =   16
         Top             =   1890
         Width           =   1395
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
         Height          =   345
         Left            =   1290
         TabIndex        =   15
         Top             =   210
         Width           =   1395
      End
      Begin VB.TextBox txtName 
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
         Height          =   345
         Left            =   1290
         TabIndex        =   14
         Top             =   630
         Width           =   1395
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "확인"
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
         Left            =   330
         TabIndex        =   12
         Top             =   3270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
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
         Left            =   1590
         TabIndex        =   11
         Top             =   3300
         Width           =   1215
      End
      Begin VB.TextBox txtSex 
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
         Height          =   345
         Left            =   1290
         TabIndex        =   10
         Top             =   1050
         Width           =   915
      End
      Begin VB.TextBox txtAge 
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
         Height          =   345
         Left            =   1290
         TabIndex        =   9
         Top             =   1470
         Width           =   915
      End
      Begin VB.CheckBox chkAllOrder 
         Caption         =   "Check1"
         Height          =   345
         Left            =   3540
         TabIndex        =   8
         Top             =   390
         Width           =   225
      End
      Begin VB.TextBox txtDate 
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
         Height          =   345
         Left            =   300
         TabIndex        =   7
         Top             =   2940
         Visible         =   0   'False
         Width           =   1965
      End
      Begin FPSpread.vaSpread vasOrder 
         Height          =   3555
         Left            =   3000
         TabIndex        =   13
         Top             =   240
         Width           =   4455
         _Version        =   393216
         _ExtentX        =   7858
         _ExtentY        =   6271
         _StockProps     =   64
         ColHeaderDisplay=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         MaxRows         =   100
         ScrollBars      =   2
         SpreadDesigner  =   "frmPatSear.frx":0000
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '투명
         Caption         =   "SampleID"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1950
         Width           =   1005
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '투명
         Caption         =   "환자번호"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '투명
         Caption         =   "환자이름"
         BeginProperty Font 
            Name            =   "굴림체"
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
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "성별"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1110
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '투명
         Caption         =   "나이"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1530
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdWorkList 
      Caption         =   "WorkList"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5310
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   7890
      Width           =   1395
   End
   Begin MSComCtl2.MonthView monvCal 
      Height          =   2220
      Left            =   5520
      TabIndex        =   3
      Top             =   1410
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   35651585
      CurrentDate     =   36878
   End
   Begin VB.CheckBox chkAll 
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   660
      TabIndex        =   2
      Top             =   840
      Width           =   165
   End
   Begin FPSpread.vaSpread vasList 
      Height          =   7065
      Left            =   90
      TabIndex        =   1
      Top             =   750
      Width           =   9285
      _Version        =   393216
      _ExtentX        =   16378
      _ExtentY        =   12462
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   14
      MaxRows         =   100
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   15526606
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmPatSear.frx":10AD
   End
   Begin VB.CommandButton btnClear 
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
      Height          =   495
      Left            =   7425
      TabIndex        =   48
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "조  회"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7170
      TabIndex        =   47
      Top             =   -1140
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   1440
      TabIndex        =   42
      Top             =   7800
      Width           =   3555
      Begin VB.TextBox txtPos 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
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
         TabIndex        =   44
         Text            =   "1"
         Top             =   195
         Width           =   675
      End
      Begin VB.TextBox txtRack 
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "굴림체"
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
         TabIndex        =   43
         Text            =   "0"
         Top             =   195
         Width           =   675
      End
      Begin Threed.SSPanel SSPanel4 
         Height          =   315
         Left            =   90
         TabIndex        =   45
         Top             =   180
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Disk"
         BackColor       =   15526606
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.76
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
         TabIndex        =   46
         Top             =   180
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   556
         _StockProps     =   15
         Caption         =   "Pos"
         BackColor       =   15526606
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9.76
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   675
      Left            =   90
      TabIndex        =   39
      Top             =   750
      Visible         =   0   'False
      Width           =   8865
      _Version        =   65536
      _ExtentX        =   15637
      _ExtentY        =   1191
      _StockProps     =   15
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.CommandButton btnSch 
         Caption         =   "Command3"
         Height          =   270
         Left            =   5025
         TabIndex        =   49
         Top             =   255
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBarCode 
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
         Left            =   1500
         TabIndex        =   41
         Top             =   150
         Width           =   2385
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검체번호"
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
         Left            =   450
         TabIndex        =   40
         Top             =   210
         Width           =   900
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   675
      Left            =   90
      TabIndex        =   32
      Top             =   60
      Width           =   9285
      _Version        =   65536
      _ExtentX        =   16378
      _ExtentY        =   1191
      _StockProps     =   15
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.01
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.Frame Frame1 
         Height          =   435
         Left            =   4440
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
         Begin VB.OptionButton optGubun 
            Caption         =   "모두"
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   55
            Top             =   150
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.OptionButton optGubun 
            Caption         =   "검진"
            Height          =   225
            Index           =   1
            Left            =   1620
            TabIndex        =   54
            Top             =   150
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.OptionButton optGubun 
            Caption         =   "진료"
            Height          =   225
            Index           =   0
            Left            =   780
            TabIndex        =   53
            Top             =   150
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   765
         End
      End
      Begin VB.OptionButton optState 
         BackColor       =   &H00E0E0E0&
         Caption         =   "전체"
         Height          =   315
         Index           =   1
         Left            =   6600
         TabIndex        =   51
         Top             =   210
         Width           =   1095
      End
      Begin VB.OptionButton optState 
         BackColor       =   &H00E0E0E0&
         Caption         =   "접수"
         Height          =   315
         Index           =   0
         Left            =   5790
         TabIndex        =   50
         Top             =   210
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2460
         TabIndex        =   38
         Top             =   210
         Width           =   255
      End
      Begin VB.CommandButton cmdCalendar 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4380
         TabIndex        =   37
         Top             =   210
         Width           =   255
      End
      Begin VB.TextBox dtpSDate 
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
         TabIndex        =   34
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox dtpEDate 
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
         Left            =   3120
         TabIndex        =   33
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "~"
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
         Left            =   2850
         TabIndex        =   36
         Top             =   270
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "접수일자"
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
         Left            =   150
         TabIndex        =   35
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdUp 
      Height          =   525
      Left            =   150
      Picture         =   "frmPatSear.frx":266A
      Style           =   1  '그래픽
      TabIndex        =   28
      Top             =   7890
      Width           =   615
   End
   Begin VB.CommandButton cmdDown 
      Height          =   525
      Left            =   780
      Picture         =   "frmPatSear.frx":2799
      Style           =   1  '그래픽
      TabIndex        =   27
      Top             =   7890
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   525
      Left            =   8085
      Style           =   1  '그래픽
      TabIndex        =   26
      Top             =   7890
      Width           =   1305
   End
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order 전송"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6750
      Style           =   1  '그래픽
      TabIndex        =   24
      Top             =   7890
      Width           =   1305
   End
   Begin VB.TextBox txtStart 
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3270
      TabIndex        =   23
      Text            =   "1"
      Top             =   7980
      Visible         =   0   'False
      Width           =   885
   End
   Begin FPSpread.vaSpread vasPrint 
      Height          =   1560
      Left            =   3420
      TabIndex        =   5
      Top             =   1500
      Visible         =   0   'False
      Width           =   5190
      _Version        =   393216
      _ExtentX        =   9155
      _ExtentY        =   2752
      _StockProps     =   64
      ColHeaderDisplay=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   8
      MaxRows         =   100
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   15526606
      ShadowDark      =   13815180
      SpreadDesigner  =   "frmPatSear.frx":28CB
   End
   Begin FPSpread.vaSpread vasCode 
      Height          =   1575
      Left            =   5460
      TabIndex        =   4
      Top             =   1230
      Visible         =   0   'False
      Width           =   3045
      _Version        =   393216
      _ExtentX        =   5371
      _ExtentY        =   2778
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
      SpreadDesigner  =   "frmPatSear.frx":3A2F
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "출력"
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6780
      Style           =   1  '그래픽
      TabIndex        =   25
      Top             =   7890
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "시작 S.No : "
      BeginProperty Font 
         Name            =   "굴림체"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1860
      TabIndex        =   29
      Top             =   8055
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '투명
      Caption         =   "결과완료 : 빨간색, 미완료 : 검정색"
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   180
      TabIndex        =   22
      Top             =   7500
      Visible         =   0   'False
      Width           =   3675
   End
End
Attribute VB_Name = "frmPatSear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iIndex As Integer

Public glRow As Long
Public gOCnt As Integer
Public gCount As String

Private Sub btnClear_Click()
    ClearSpread vasList
    
End Sub

Private Sub btnSch_Click()
    Dim sSch1, sSch2 As String
    Dim iRow As Integer
    Dim I As Integer
    Dim sCnt As String
    Dim sExamCode As String
    Dim sExamName As String
        
    'vasList.MaxRows = 100

    '체크, Rack, Pos, SampleNo, 환자번호, 환자이름, 성별, 나이, 주민번호, 접수일자
    '검사상태
    sSch1 = Format(dtpSDate.Text, "yymmdd") & "0001"
    sSch2 = Format(dtpEDate.Text, "yymmdd") & "9999"
        
    SQL = "SELECT a.PTNO, " & vbCrLf
    SQL = SQL & " a.SNAME, a.SEX, a.AGE, '', " & vbCrLf
    SQL = SQL & " '20' || substr(a.SPECNO, 1, 6), substr(a.SPECNO, 7, 4), a.SPECNO, count(SUBCODE) " & vbCrLf
    SQL = SQL & "From TWEXAM_SPECMST a, TWEXAM_RESULTC b " & vbCrLf
    SQL = SQL & "WHERE a.SPECNO = '" & Trim(txtBarCode) & "' " & vbCrLf
    SQL = SQL & "  AND b.SPECNO = a.SPECNO " & vbCrLf
    SQL = SQL & "  AND b.SUBCODE In (" & gAllExam & ") " & vbCrLf
    SQL = SQL & "  AND b.STATUS in ('2','3') " & vbCrLf
    SQL = SQL & "Group by a.PTNO, " & vbCrLf
    SQL = SQL & " a.SNAME, a.SEX, a.AGE, '', a.BDATE, " & vbCrLf
    SQL = SQL & " '20' || substr(a.SPECNO, 1, 6), substr(a.SPECNO, 7, 4), a.SPECNO "
    res = db_select_Vas(gServer, SQL, vasList, vasList.DataRowCnt + 1, 4)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    'vasSort vasList, 11
    
    For iRow = 1 To vasList.DataRowCnt
        sExamCode = ""
        sExamName = ""
        ClearSpread vasOrder
        
        SQL = "SELECT SUBCODE " & vbCrLf
        SQL = SQL & "From TWEXAM_RESULTC  " & vbCrLf
        SQL = SQL & "WHERE SPECNO = '" & Trim(GetText(vasList, iRow, 11)) & "' " & vbCrLf
        SQL = SQL & "  AND SUBCODE In (" & gAllExam & ") " & vbCrLf
        SQL = SQL & "  AND STATUS in ('2','3') "
        res = db_select_Vas(gServer, SQL, vasOrder)
        vasSort vasOrder, 1
        
        For I = 1 To vasOrder.DataRowCnt
            sExamCode = sExamCode & "'" & Trim(GetText(vasOrder, I, 1)) & "',"
        Next I
        If Len(sExamCode) > 0 Then
            sExamCode = Left(sExamCode, Len(sExamCode) - 1)
        End If
        ClearSpread vasOrder
        SQL = "Select examname From equipexam" & vbCrLf & _
              " Where Equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examcode in (" & sExamCode & ") "
        res = db_select_Vas(gLocal, SQL, vasOrder)
        For I = 1 To vasOrder.DataRowCnt
            sExamName = sExamName & Trim(GetText(vasOrder, I, 1)) & "/"
        Next I
        If Len(sExamName) > 0 Then
            sExamName = Left(sExamName, Len(sExamName) - 1)
            
            vasList.Row = iRow
            vasList.Col = 1
            vasList.Value = 1
        End If
        vasList.SetText 12, iRow, sExamName
        
        vasList.Row = iRow
        vasList.Col = 2
        vasList.TypeComboBoxCurSel = 0
        
        SQL = "select state, SEQNO from Worklist " & vbCrLf & _
              "WHERE examdate = '" & Format(CDate(frmInterface.txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
              "  AND SampleID = '" & Trim(GetText(vasList, iRow, 11)) & "' "
        res = db_select_Col(gLocal, SQL)
        vasList.SetText 3, iRow, Trim(gReadBuf(1))
        Select Case Trim(gReadBuf(0))
        Case "A"
            SetBackColor vasList, iRow, iRow, 5, 5, 255, 255, 112
        Case "B", "C"
            SetBackColor vasList, iRow, iRow, 5, 5, 202, 255, 112
        Case Else
            SetBackColor vasList, iRow, iRow, 5, 5, 255, 255, 255
        End Select
    Next iRow
    
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 13.3

End Sub

Private Sub chkAll_Click()
    Dim iRow As Integer
    
    If chkAll.Value = 1 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 1
        Next iRow
    ElseIf chkAll.Value = 0 Then
        For iRow = 1 To vasList.DataRowCnt
            vasList.Row = iRow
            vasList.Col = 1
            
            vasList.Value = 0
        Next iRow
    End If
End Sub

Private Sub chkAllOrder_Click()
    If chkAllOrder.Value = 1 Then
        vasOrder.Row = -1
        vasOrder.Col = 1
        vasOrder.Value = 1
    Else
        vasOrder.Row = -1
        vasOrder.Col = 1
        vasOrder.Value = 0
    End If
End Sub

Private Sub cmdCalendar_Click(Index As Integer)
    iIndex = Index
    If Index = 0 Then
        monvCal.Left = dtpSDate.Left
        monvCal.Top = 570
        monvCal.Visible = True
        
        monvCal.Value = dtpSDate.Text
    ElseIf Index = 1 Then
        monvCal.Left = dtpEDate.Left
        monvCal.Top = 570
        monvCal.Visible = True
        
        monvCal.Value = dtpEDate.Text
    End If
    'monvCal.Visible = True
End Sub

Private Sub cmdClose_Click()
'    txtDate.Text = ""
'    txtPID.Text = ""
'    txtName.Text = ""
'    txtSex.Text = ""
'    txtAge.Text = ""
'
'    ClearSpread vasOrder
'
    sspOrder.Visible = False
End Sub

Private Sub cmdDown_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, vasList.MaxCols, lRow, 1, lRow + 1
    vasActiveCell vasList, lRow + 1, 2
    vasList_Click 2, lRow + 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'Local에 환자에 대한 검사항목 저장하기
Dim sCnt As String
Dim iRow As Integer
Dim sExamCode As String
Dim sEquipCode As String
Dim sAge As String
Dim I As Integer

    sCnt = ""
    
    SQL = " Select count(*) From pat_res " & vbCrLf & _
          " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
          " And equipno = '" & gEquip & "' " & vbCrLf & _
          " And barcode = '" & Trim(txtNo) & "' " & vbCrLf & _
          " And sendflag = 'O' "
    res = db_select_Var(gLocal, SQL, sCnt)
    
    If sCnt = "" Then
        sCnt = "0"
    End If
    
    If txtAge.Text = "" Then
        txtAge.Text = "0"
    Else
        sAge = Trim(txtAge.Text)
    End If
    
    If sCnt > 0 Then
            SQL = " Delete From pat_res " & vbCrLf & _
                  " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
                  " And equipno = '" & gEquip & "' " & vbCrLf & _
                  " And barcode = '" & Trim(txtNo.Text) & "' " & vbCrLf & _
                  " And sendflag = 'O' "
            res = SendQuery(gLocal, SQL)
            
            If res = -1 Then
                SaveQuery SQL
            End If
    End If
    
    For iRow = 1 To vasOrder.DataRowCnt
        vasOrder.Row = iRow
        vasOrder.Col = 1
        
        If vasOrder.Value = 1 Then
            sExamCode = Trim(GetText(vasOrder, iRow, 2))
            sEquipCode = GetEquip_ExamCode(sExamCode)

            SQL = " Insert Into pat_res(examdate, equipno, barcode, equipcode,  " & vbCrLf & _
                  " examcode, pid, pname, psex, page, resdate, sendflag)  " & vbCrLf & _
                  " Values ( '" & Trim(txtDate) & "', '" & gEquip & "', '" & Trim(txtNo.Text) & "' , '" & Trim(sEquipCode) & "', " & vbCrLf & _
                  " '" & sExamCode & "', '" & Trim(txtPID.Text) & "', " & vbCrLf & _
                  " '" & Trim(txtName.Text) & "', '" & Trim(txtSex.Text) & "', " & sAge & ", " & vbCrLf & _
                  " '" & Trim(GetDateFull) & "', 'O') "
            res = SendQuery(gLocal, SQL)
            
            If res = -1 Then
                SaveQuery SQL
            End If
        ElseIf vasOrder.Value = 0 Then
            If sCnt = 0 Then
            
            ElseIf sCnt > 0 Then
                sExamCode = Trim(GetText(vasOrder, iRow, 2))
                
                SQL = " Delete From pat_res " & vbCrLf & _
                      " Where examdate = '" & Trim(txtDate) & "' " & vbCrLf & _
                      " And equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(txtNo.Text) & "' " & vbCrLf & _
                      " And examcode = '" & sExamCode & "' "
                res = SendQuery(gLocal, SQL)
                
                If res = -1 Then
                    SaveQuery SQL
                End If
            End If
        End If
    Next iRow
    
    sspOrder.Visible = False
End Sub

'
'Private Sub cmdOrder_Click()
'    Dim llRow_Order As Long
'    Dim iRow As Integer
'    Dim jRow As Integer
'    Dim I As Integer
'    Dim iCnt As Integer
'
'    Dim sEquipCode As String
'    Dim sOrderCode As String
'    Dim sOrder As String
'
'    Dim sID As String
'
'    Dim lsCurDate As String
'    Dim lsSampleNo As String
'    Dim lsType As String
'    Dim lsTypeSelect As Integer
'
'    If IsNumeric(txtRack) = False Or IsNumeric(txtPos) = False Then
'        MsgBox "Rack, Pos을 확인하세요!", vbCritical, "알림"
'        Exit Sub
'    End If
'
''    If IsNumeric(txtStart) Then
''        lsSampleNo = Trim(txtStart)
''    Else
''        lsSampleNo = "1"
''    End If
'
'    lsCurDate = Format(Date, "yyyymmdd") & Format(Time, "hhnnss")
'
''    ClearSpread frmInterface.vasOrder
'
'    llRow_Order = 1
'
'    For iRow = 1 To vasList.DataRowCnt
'        If Trim(GetText(vasList, iRow, 3)) <> "" Then
'            SetText vasList, Format(Trim(GetText(vasList, iRow, 3)), "0#"), iRow, 3
'        End If
'    Next iRow
'
'    vasSort vasList, 3
'
'    For iRow = 1 To vasList.DataRowCnt
'        vasList.Row = iRow
'        vasList.Col = 1
'
'        If vasList.Value = 1 Then
'            '처방가져오기
'            sOrderCode = ""
'
'            vasList.SetText 3, iRow, txtPos
'
'            txtPos = CStr(CInt(txtPos) + 1)
'
'            ClearSpread vasCode
'
'            sID = Trim(GetText(vasList, iRow, 10))     '검체번호
'
''            If Trim(GetText(vasList, iRow, 3)) = "" Then
''                SetText vasList, txtPos, iRow, 3
''            End If
''
''            lsSampleNo = CLng(lsSampleNo) + 1
''            txtStart = lsSampleNo
'
'            frmInterface.vasOrder.SetText 1, llRow_Order, sID
'            frmInterface.vasOrder.SetText 2, llRow_Order, Trim(txtRack)
'            'frmInterface.vasOrder.SetText 3, llRow_Order, Trim(txtPos)
'            frmInterface.vasOrder.SetText 3, llRow_Order, Trim(GetText(vasList, iRow, 3))
'            frmInterface.vasOrder.SetText 4, llRow_Order, ""
'
'            llRow_Order = llRow_Order + 1
'            If llRow_Order > frmInterface.vasOrder.MaxRows Then
'                frmInterface.vasOrder.MaxRows = llRow_Order
'            End If
'
''            If IsNumeric(txtPos) Then
''                txtPos = CInt(txtPos) + 1
''            End If
'        End If
'    Next iRow
'
'    'WorkList 전송
'    cmdWorkList_Click
'
''    gRecodeType = "Q"
''
''    comSend = "stENQ"
'
'    If frmInterface.vasOrder.DataRowCnt > 0 Then
'        gOrderMessage = Trim(GetText(frmInterface.vasOrder, 1, 1))
'        gRack = Trim(GetText(frmInterface.vasOrder, 1, 2))
'        gPos = Trim(GetText(frmInterface.vasOrder, 1, 3))
'        gSampleNo = ""
'
'        gOrderCnt = 0
'
'        gPreMsg = chrENQ
'        Save_Raw_Data "[Tx]" & gPreMsg
'        frmInterface.MSComm1.Output = gPreMsg
'    End If
'
'    Unload Me
'End Sub

Private Sub cmdPrint_Click()
Dim iRow As Integer
Dim j As Integer

Dim sCurDate As String
Dim sSerDate As String
Dim sHead As String
Dim sFoot As String
    
    ClearSpread vasPrint

    j = 1

    If optGubun(1).Value = True Then
        vasPrint.RowHeight(-1) = 39.2
    Else
        vasPrint.RowHeight(-1) = 25.9
    End If
    
    For iRow = 1 To vasList.DataRowCnt
        vasList.Row = iRow
        vasList.Col = 1

        If vasList.Value = 1 Then
            SetText vasPrint, Trim(GetText(vasList, iRow, 11)), j, 1     '검체번호
            SetText vasPrint, Trim(GetText(vasList, iRow, 4)), j, 2     '환자번호
            SetText vasPrint, Trim(GetText(vasList, iRow, 5)), j, 3     '환자이름

            SetText vasPrint, Trim(GetText(vasList, iRow, 6)), j, 4     '성별
            SetText vasPrint, Trim(GetText(vasList, iRow, 7)), j, 5     '나이
            'SetText vasPrint, Trim(GetText(vasList, iRow, 7)), j, 6     '주민번호
            SetText vasPrint, Trim(GetText(vasList, iRow, 9)), j, 7     '처방일자
            SetText vasPrint, Trim(GetText(vasList, iRow, 12)), j, 8     '처방일자
            
            j = j + 1
        End If
    Next iRow
    
    If vasPrint.DataRowCnt < 1 Then
        MsgBox "출력할 자료가 없습니다.", , "알 림"
        Exit Sub
    End If
    
    sCurDate = GetDateFull
    
    sSerDate = Trim(dtpSDate.Text) & " - " & Trim(dtpEDate.Text)
    
    '2004/08/11 이상은 - 세로방향에서 가로방향으로 수정
    vasPrint.PrintOrientation = 1   ' SS_PRINTORIENT_PORTRAIT
    vasPrint.PrintAbortMsg = "인쇄중 입니다 ..."
    vasPrint.PrintJobName = "Cobas E411 WorkList 출력"
    

    sHead = "/fn""궁서체"" /fz""12"" /fb1 /fi0 /fu0 " & "/c" & "▣ Modular WorkList ▣" & "/n/n " & _
            "/fn""굴림체"" /fz""10"" /fb0 /fi0 /fu0 " & "/c" & "처방일자 : " & dtpSDate & " ~ " & dtpEDate
    If optGubun(0).Value = True Then
        sHead = sHead & " (진료)" & "/n/n"
    ElseIf optGubun(1).Value = True Then
        sHead = sHead & " (검진)" & "/n/n"
    End If

    sFoot = "/fn""굴림체"" /fz""10"" /fb1 /fi0 /fu0 " & "/l" & sCurDate & "/fn""궁서체"" /fz""11"" /fb1 /fi0 /fu0 /r" & " 굿모닝병원 검사실"
    
    vasPrint.PrintHeader = sHead
    vasPrint.PrintFooter = sFoot

    vasPrint.PrintMarginTop = 680
    vasPrint.PrintMarginBottom = 680
'현재 SS가 비대칭으로 출력함
'    vaslist.PrintMarginLeft = 720
    vasPrint.PrintMarginLeft = 0
    vasPrint.PrintMarginRight = 0
    
    vasPrint.PrintColor = True
    vasPrint.PrintGrid = True
    
'Set printing range
    vasPrint.PrintType = 0  'SS_PRINT_ALL(default)

    vasPrint.PrintShadows = True

    vasPrint.Action = 13 'SS_ACTION_PRINT
End Sub

Private Sub cmdSearch_1_Click()
    Dim sSch1, sSch2 As String
    Dim iRow As Integer
    Dim sCnt As String
    
    ClearSpread vasList
    
    vasList.MaxRows = 100
    
    
    '체크, Rack, Pos, SampleNo, 환자번호, 환자이름, 성별, 나이, 주민번호, 접수일자
    '검사상태
    sSch1 = Format(dtpSDate.Text, "yyyy-mm-dd")
    sSch2 = Format(dtpEDate.Text, "yyyy-mm-dd")
    
    SQL = " Select max(a.DR_CHART), b.PE_SUJINJA, '', '', b.PE_JUMIN, a.DR_DATE, '', '' " & vbCrLf & _
          " From DEPARTDAT a, PERSON b " & vbCrLf & _
          " Where a.DR_DATE between '" & sSch1 & "' and '" & sSch2 & "' " & vbCrLf & _
          " And a.DR_CODE in (" & gAllExam & ") " & vbCrLf & _
          " And a.DR_CHART = b.PE_CHART "
          
'    If optState(0).Value = True Then        '접수
'        SQL = SQL & vbCrLf & _
'              " And c.GD_RESULT = ''  "
'    ElseIf optState(1).Value = True Then    '결과
'        SQL = SQL & vbCrLf & _
'              " And c.GD_RESULT <> '' "
'    ElseIf optState(2).Value = True Then
'    End If
        
        SQL = SQL & vbCrLf & _
              " Group by b.PE_SUJINJA, b.PE_JUMIN, a.DR_DATE " & vbCrLf & _
              " Order by 1 "

    res = db_select_Vas(gServer, SQL, vasList, 1, 5)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasList.MaxRows = vasList.DataRowCnt
    
    For iRow = 1 To vasList.DataRowCnt
        CalAgeSex Trim(GetText(vasList, iRow, 9)), Format(dtpSDate.Text, "yyyy/mm/dd")
        If gPatGen.Age = "" Then
            gPatGen.Age = 0
        End If
        SetText vasList, gPatGen.Sex, iRow, 7
        SetText vasList, gPatGen.Age, iRow, 8
        
        sCnt = ""
        
        SQL = " Select count(GD_CODE) From GUMSADAT " & vbCrLf & _
              " Where GD_DATE = '" & Trim(GetText(vasList, iRow, 10)) & "' " & vbCrLf & _
              " And GD_CHART = '" & Trim(GetText(vasList, iRow, 5)) & "' " & vbCrLf & _
              " And GD_CODE in (" & gAllExam & ") "
        res = db_select_Var(gServer, SQL, sCnt)
        
        If sCnt = "" Then
            sCnt = "0"
        End If
        
        If sCnt = "0" Then
            SetForeColor vasList, iRow, iRow, 0, 0, 0
        ElseIf CInt(sCnt) > 0 Then
            SetForeColor vasList, iRow, iRow, 250, 0, 0
        End If
    Next iRow

End Sub

Private Sub cmdSearch_Click()
    Dim sSch1, sSch2 As String
    Dim iRow As Integer
    Dim I As Integer
    Dim sCnt As String
    Dim sExamCode As String
    Dim sExamName As String
    
    ClearSpread vasList
    
    '체크, Type, SampleNo, 환자번호, 환자이름, 성별, 나이, 주민번호, 접수일자, 접수번호, SampleID
    '0 : 접수 1 : 예비 2 : 일부 3 : 최종 4 : 수정
    
'    sSch1 = SeperatorCls(Trim(dtpSDate.Text)) & "0000"
'    sSch2 = SeperatorCls(Trim(dtpEDate.Text)) & "2359"
'
'    SQL = "SELECT a.pt_no, " & vbCrLf
'    SQL = SQL & " a.pt_nm, '', '', a.ssn_1 || a.ssn_2, " & vbCrLf
'    SQL = SQL & " b.acp_dt, '', b.smp_no, count(b.cd) " & vbCrLf
'    SQL = SQL & "From pmcptbsm a, scrprexh b " & vbCrLf
'    SQL = SQL & "WHERE a.hos_org_no = '38203111' " & vbCrLf
'    SQL = SQL & "  AND b.acp_dt >= '" & sSch1 & "' " & vbCrLf
'    SQL = SQL & "  AND b.acp_dt <= '" & sSch2 & "' " & vbCrLf
'    SQL = SQL & "  AND a.hos_org_no = b.hos_org_no " & vbCrLf
'    SQL = SQL & "  AND a.pt_no = b.pt_no " & vbCrLf
'    SQL = SQL & "  AND b.cd In (" & gAllExam & ") " & vbCrLf
'
'    If optState(0).Value = True Then
'        SQL = SQL & "  AND b.smp_stus in ('0') " & vbCrLf
'    Else
'        SQL = SQL & "  AND b.smp_stus not in ('0','3') " & vbCrLf
'    End If
'
'    SQL = SQL & "Group by a.pt_no, " & vbCrLf
'    SQL = SQL & " a.pt_nm, a.ssn_1 || a.ssn_2, b.acp_dt, b.smp_no "
    
    sSch1 = SeperatorCls(Trim(dtpSDate.Text))
    sSch2 = SeperatorCls(Trim(dtpEDate.Text))
    
    SQL = "Select b.PbsChtNum, b.PbsPatNam,b.PbsSexTyp,b.PbsResNum,'',a.Rstodrdte,a.RstLabNum "
    SQL = SQL & vbCrLf & " from Rstinf a, Pbsinf b, LabInf c "
    SQL = SQL & vbCrLf & " where b.pbschtnum = c.labchtnum "
    SQL = SQL & vbCrLf & " and  a.rstlabnum = c.lablabnum "
    SQL = SQL & vbCrLf & " and  a.rstodrdte BETWEEN '" & sSch1 & "' AND '" & sSch2 & "' "
    SQL = SQL & vbCrLf & "   and a.RstOdrCod In (" & gAllExam & ") "
    If optState(0).Value = True Then
        SQL = SQL & vbCrLf & "   and a.RstRstVal= '' "
    End If
'    SQL = SQL & vbCrLf & "   and b.workarea = a.workarea "
'    SQL = SQL & vbCrLf & "   and b.accdt = a.accdt "
'    SQL = SQL & vbCrLf & "   and b.accseq = a.accseq "
'    SQL = SQL & vbCrLf & "   AND b.testcd In (" & gAllExam & ") "
'    SQL = SQL & vbCrLf & "   and c.ptnt_no = a.ptid "
'
    SQL = SQL & vbCrLf & " group by b.PbsChtNum, b.PbsPatNam, b.PbsResNum, b.PbsSexTyp, a.Rstodrdte,a.RstLabNum "
    
    res = db_select_Vas(gServer, SQL, vasList, 1, 4)
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    

    For iRow = 1 To vasList.DataRowCnt
        '성별/나이
        CalAgeSex Trim(GetText(vasList, iRow, 7)), Format(CDate(GetDateFull), "YYYY/MM/DD")
        SetText vasList, gPatGen.Sex, iRow, 6
        SetText vasList, gPatGen.Age, iRow, 7
'        If IsNumeric(Trim(GetText(vasList, iRow, 10))) Then
'            vasList.SetText 11, iRow, Trim(GetText(vasList, iRow, 9)) & Format(CCur(Trim(GetText(vasList, iRow, 10))), "00000")
'        End If

        sExamCode = ""
        sExamName = ""
        ClearSpread vasOrder

        SQL = "SELECT Rstodrcod " & vbCrLf
        SQL = SQL & "From rstinf  " & vbCrLf
        SQL = SQL & "WHERE rstlabnum = '" & Trim(GetText(vasList, iRow, 4)) & "'  " & vbCrLf
        SQL = SQL & "  AND rstodrcod  in (" & gAllExam & ") " '& vbCrLf
        'SQL = SQL & "  AND accseq  = " & Trim(GetText(vasList, iRow, 10)) & vbCrLf
        'SQL = SQL & "  AND testcd In (" & gAllExam & ") " & vbCrLf

        res = db_select_Vas(gServer, SQL, vasOrder)
        vasSort vasOrder, 1

        For I = 1 To vasOrder.DataRowCnt
            sExamCode = sExamCode & "'" & Trim(GetText(vasOrder, I, 1)) & "',"
        Next I
        If Len(sExamCode) > 0 Then
            sExamCode = Left(sExamCode, Len(sExamCode) - 1)
        End If
        ClearSpread vasOrder
        SQL = "Select examname From equipexam" & vbCrLf & _
              " Where Equipno = '" & gEquip & "' " & vbCrLf & _
              "  and examcode in (" & sExamCode & ") "
        res = db_select_Vas(gLocal, SQL, vasOrder)
        For I = 1 To vasOrder.DataRowCnt
            sExamName = sExamName & Trim(GetText(vasOrder, I, 1)) & "/"
        Next I
        If Len(sExamName) > 0 Then
            sExamName = Left(sExamName, Len(sExamName) - 1)
        End If
        vasList.SetText 12, iRow, sExamName

        vasList.Row = iRow
        vasList.Col = 2
        vasList.TypeComboBoxCurSel = 0
'
'        SQL = "select state, SEQNO from Worklist " & vbCrLf & _
'              "WHERE examdate = '" & Format(CDate(frmInterface.txtToday.Text), "yyyymmdd") & "' " & vbCrLf & _
'              "  AND SampleID = '" & Trim(GetText(vasList, iRow, 11)) & "' "
'        res = db_select_Col(gLocal, SQL)
'        vasList.SetText 3, iRow, Trim(gReadBuf(1))
'        Select Case Trim(gReadBuf(0))
'        Case "A"
'            SetBackColor vasList, iRow, iRow, 5, 5, 255, 255, 112
'        Case "B", "C"
'            SetBackColor vasList, iRow, iRow, 5, 5, 202, 255, 112
'        Case Else
'            SetBackColor vasList, iRow, iRow, 5, 5, 255, 255, 255
'        End Select
    Next iRow
    
    vasSort vasList, 11
    
    vasList.MaxRows = vasList.DataRowCnt
    vasList.RowHeight(-1) = 13.3
End Sub

Private Sub cmdUp_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    vasList.SwapRange 1, lRow, vasList.MaxCols, lRow, 1, lRow - 1
    vasActiveCell vasList, lRow - 1, 2
    vasList_Click 2, lRow - 1
End Sub

Private Sub cmdWorkList_Click()
    Dim lRow As Long
    Dim lCol As Long
    Dim lDestRow As Long
    
    'If frmInterface.vasOrder.DataRowCnt > 0 Then
    
    'Else
        lDestRow = frmInterface.vasID.DataRowCnt + 1
    
        If frmInterface.vasID.MaxRows < lDestRow Then
            frmInterface.vasID.MaxRows = lDestRow
        End If
        
        For lRow = 1 To vasList.DataRowCnt
            vasList.Row = lRow
            vasList.Col = 1
            If vasList.Value = 1 And Trim(GetText(vasList, lRow, 4)) <> "" Then
                For lCol = 2 To 11
                    Select Case lCol
                    Case 3
                        If IsNumeric(Trim(GetText(vasList, lRow, lCol))) = True Then
                            SetText frmInterface.vasID, CInt(Trim(GetText(vasList, lRow, lCol))), lDestRow, 4
                        Else
                            SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 4
                        End If
                    Case 4, 5, 6, 7
                        SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 2
                    Case 8      '주민번호
                        SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 4
                    Case 9      '접수일자
                        SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 4
                    Case 10     '검체번호
                        SetText frmInterface.vasID, Trim(GetText(vasList, lRow, lCol)), lDestRow, 5
                    End Select
                    SetText frmInterface.vasID, "", lDestRow, 2
                    SetText frmInterface.vasID, "", lDestRow, 3
                    SetText frmInterface.vasID, "오더", lDestRow, 11
                Next lCol
                
                SQL = "delete from worklist where ExamDate = '" & Format(CDate(frmInterface.txtToday), "yyyymmdd") & "' and SampleID = '" & Trim(GetText(vasList, lRow, 11)) & "'"
                res = SendQuery(gLocal, SQL)
                
                If Not IsNumeric(GetText(vasList, lRow, 9)) Then
                    vasList.SetText 9, lRow, "0"
                End If
'                SQL = "Insert Into worklist ( ExamDate, SeqNo, SampleID, PID, PName, PSex, PAge, ExamName,State )"
'                SQL = SQL & vbCrLf & "Values ('" & Format(CDate(frmInterface.txtToday), "yyyymmdd") & "', "
'                SQL = SQL & vbCrLf & Trim(GetText(vasList, lRow, 3)) & ", "
'                SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, lRow, 11)) & "', "
'                SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, lRow, 4)) & "', "
'                SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, lRow, 5)) & "', "
'                SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, lRow, 6)) & "', "
'                SQL = SQL & vbCrLf & "'" & Trim(GetText(vasList, lRow, 7)) & "', "
'                SQL = SQL & vbCrLf & "'', 'A') "
'                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    Save_Raw_Data "[SQL : " & res & "]" & SQL
                End If
                
                lDestRow = lDestRow + 1
            End If
        Next lRow
        
        'ChkAll.Value = 0
    'End If
End Sub

Private Sub Command1_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    If lRow = 1 Then Exit Sub
    
    lRow = lRow - 1
    
    vasActiveCell vasList, lRow, 5
    
    vasList_DblClick 5, lRow
    
End Sub

Private Sub Command2_Click()
    Dim lRow As Long
    
    lRow = vasList.ActiveRow
    
    If lRow = vasList.DataRowCnt Then Exit Sub
    
    lRow = lRow + 1
    
    vasActiveCell vasList, lRow, 5
    
    vasList_DblClick 5, lRow
End Sub

Private Sub Form_Activate()
    'dtpSDate.SetFocus
    vasActiveCell vasList, 1, 2
End Sub

Private Sub Form_Load()

    dtpSDate.Text = Format(DateAdd("y", CDate(GetDateFull), -2), "yyyy/mm/dd")
    dtpEDate.Text = Format(CDate(GetDateFull), "YYYY/MM/DD")
    
    ClearSpread vasList
    
    chkAll.Value = 0
    
    cmdSearch_Click
    
End Sub

Private Sub monvCal_DateClick(ByVal DateClicked As Date)
    If iIndex = 0 Then
        dtpSDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
    Else
        dtpEDate.Text = Trim(Format(DateClicked, "yyyy-mm-dd"))
    End If
    monvCal.Visible = False
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtBarCode_GotFocus()
    SelectFocus txtBarCode
End Sub

Private Sub txtBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txtBarCode) <> 10 Then
            txtBarCode.SetFocus
            Exit Sub
        End If
        btnSch_Click
        txtBarCode = ""
    End If
End Sub

Private Sub vasList_Click(ByVal Col As Long, ByVal Row As Long)
    If sspOrder.Visible = True Then sspOrder.Visible = False
    
    If Row = 0 Then
        vasSort vasList, Col
    End If
    
    If Row < 0 Or Row > vasList.DataRowCnt Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    
    If Row = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    ElseIf Row = vasList.DataRowCnt Then
        cmdUp.Enabled = True
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = True
        cmdDown.Enabled = True
    End If
End Sub

Private Sub vasList_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim lRow, lCol As Long
'    Dim lDestRow As Long
'
'    lDestRow = Form_Main.vasExam.DataRowCnt + 1
'
'    lRow = vasList.ActiveRow
'
'    For lCol = 2 To 8
'        If lCol = 8 Then        '처방일자
'            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, 8)), lDestRow, 12
'        ElseIf lCol = 2 Then    '검체번호
'            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, 2)), lDestRow, 2
'        Else
'            SetText Form_Main.vasExam, Trim(GetText(vasList, lRow, lCol)), lDestRow, lCol + 3
'        End If
'    Next lCol

'    Unload Me

'===================================================================
'2004/08/03 이상은 - 환자 더블클릭시 상세 검사항목 디스플레이 되도록
Dim sCnt As String
Dim sExamCode As String
Dim sEquipCode As String

Dim iRow As Integer
Dim jRow As Integer

    txtDate = GetText(vasList, Row, 9)
    
    txtNo = Trim(GetText(vasList, Row, 10))
    txtPID = Trim(GetText(vasList, Row, 4))
    txtName = Trim(GetText(vasList, Row, 5))
    
    txtSex = Trim(GetText(vasList, Row, 6))
    txtAge = Trim(GetText(vasList, Row, 7))
    
    chkAllOrder.Value = 0
    
    ClearSpread vasOrder
    
    '검사코드 가져오기
    
    SQL = "Select '',RstOdrCod,'' "
    SQL = SQL & vbCrLf & " from Rstinf "
    SQL = SQL & vbCrLf & " where RstLabNum = '" & txtNo & "' "
    SQL = SQL & vbCrLf & "   and RstOdrCod In (" & gAllExam & ") "

    res = db_select_Vas(gServer, SQL, vasOrder)
'    vasSort vasOrder, 2
    If res = -1 Then
        SaveQuery SQL
        Exit Sub
    End If
    
    vasOrder.MaxRows = vasOrder.DataRowCnt
    
    For jRow = 1 To vasOrder.DataRowCnt
        SQL = " select ExamName from EquipExam " & vbCrLf & _
              " where equipno = '" & gEquip & "' and ExamCode = '" & Trim(GetText(vasOrder, jRow, 2)) & "' "
        res = db_select_Col(gLocal, SQL)

        If res = 1 Then
            SetText vasOrder, Trim(gReadBuf(0)), jRow, 3
        End If
    Next jRow
    
    sspOrder.Visible = True
    
End Sub

Private Sub vasList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iRow As Integer
    
    iRow = vasList.ActiveRow
    
    If KeyCode = vbKeyDelete Then
        If iRow < 1 Or iRow > vasList.DataRowCnt Then Exit Sub
        DeleteRow vasList, iRow, iRow
    End If
End Sub


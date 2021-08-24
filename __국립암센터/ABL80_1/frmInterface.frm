VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInterface 
   BorderStyle     =   0  '없음
   Caption         =   " ABL80 인터페이스"
   ClientHeight    =   10995
   ClientLeft      =   -15
   ClientTop       =   570
   ClientWidth     =   15390
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
   Picture         =   "frmInterface.frx":1272
   ScaleHeight     =   10995
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.PictureBox Picture2 
      Align           =   1  '위 맞춤
      BackColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   15330
      TabIndex        =   61
      Top             =   0
      Width           =   15390
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   6540
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   90
         Visible         =   0   'False
         Width           =   5115
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "SEND"
         Height          =   345
         Left            =   5520
         TabIndex        =   77
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "종료"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   13860
         Style           =   1  '그래픽
         TabIndex        =   75
         Top             =   540
         Width           =   1305
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00C0FFFF&
         Caption         =   "결과저장"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   11220
         Style           =   1  '그래픽
         TabIndex        =   74
         Top             =   540
         Width           =   1305
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C0FFFF&
         Caption         =   "화면정리"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   12540
         Style           =   1  '그래픽
         TabIndex        =   73
         Top             =   540
         Width           =   1305
      End
      Begin VB.CommandButton cmdCall 
         BackColor       =   &H00C0FFFF&
         Caption         =   "결과조회"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   9900
         Style           =   1  '그래픽
         TabIndex        =   72
         Top             =   540
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancelState 
         BackColor       =   &H00C0E0FF&
         Caption         =   "검증취소"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   8310
         Style           =   1  '그래픽
         TabIndex        =   71
         Top             =   540
         Width           =   1305
      End
      Begin VB.TextBox Text_Today 
         Alignment       =   2  '가운데 맞춤
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1020
         TabIndex        =   70
         Text            =   "2002/02/18"
         Top             =   150
         Width           =   2595
      End
      Begin VB.ComboBox cmbDr 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4740
         TabIndex        =   67
         Top             =   690
         Width           =   2625
      End
      Begin VB.ComboBox cmbDept 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   11.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1020
         TabIndex        =   66
         Text            =   "마취통증의학과"
         Top             =   690
         Width           =   2595
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "처방의사"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   3840
         TabIndex        =   69
         Top             =   750
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "처방부서"
         BeginProperty Font 
            Name            =   "굴림"
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
         Left            =   60
         TabIndex        =   68
         Top             =   750
         Width           =   780
      End
      Begin VB.Label Label1 
         Appearance      =   0  '평면
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '투명
         Caption         =   "검사일자"
         BeginProperty Font 
            Name            =   "굴림"
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
         Left            =   120
         TabIndex        =   65
         Top             =   210
         Width           =   780
      End
      Begin VB.Image imgPort 
         Height          =   240
         Left            =   12270
         Picture         =   "frmInterface.frx":14F5
         Top             =   90
         Width           =   240
      End
      Begin VB.Image imgSend 
         Height          =   240
         Left            =   13425
         Picture         =   "frmInterface.frx":1A7F
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgReceive 
         Height          =   240
         Left            =   14580
         Picture         =   "frmInterface.frx":2009
         Top             =   90
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "포트"
         Height          =   195
         Index           =   0
         Left            =   11760
         TabIndex        =   64
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "송신"
         Height          =   195
         Left            =   12945
         TabIndex        =   63
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "수신"
         Height          =   195
         Left            =   14070
         TabIndex        =   62
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.Frame FrmTempBox 
      Caption         =   "TempBox"
      Height          =   2205
      Left            =   13290
      TabIndex        =   49
      Top             =   2460
      Visible         =   0   'False
      Width           =   9165
      Begin VB.Timer tmrSend 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8460
         Top             =   210
      End
      Begin VB.Timer tmrReceive 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   7980
         Top             =   210
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
         Left            =   6030
         TabIndex        =   57
         Top             =   360
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
         TabIndex        =   56
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
         TabIndex        =   55
         Top             =   1650
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Command15"
         Height          =   435
         Left            =   90
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   240
         Value           =   1  '확인
         Width           =   1065
      End
      Begin MSComctlLib.ImageList imlStatus 
         Left            =   4770
         Top             =   270
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
               Picture         =   "frmInterface.frx":2593
               Key             =   "RUN"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":2B2D
               Key             =   "NOT"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":30C7
               Key             =   "STOP"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":3661
               Key             =   "LST"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":3EF3
               Key             =   "ITM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":404D
               Key             =   "ERR"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmInterface.frx":41A7
               Key             =   "NOF"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  '투명
         Caption         =   "사용자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   76
         Top             =   1830
         Width           =   1185
      End
   End
   Begin VB.Frame FrmUseControl 
      Caption         =   "UseControl"
      Height          =   975
      Left            =   13860
      TabIndex        =   48
      Top             =   6300
      Visible         =   0   'False
      Width           =   7095
      Begin VB.TextBox Text4 
         Height          =   675
         Left            =   1050
         TabIndex        =   60
         Top             =   150
         Visible         =   0   'False
         Width           =   4125
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         Height          =   435
         Left            =   5460
         TabIndex        =   59
         Top             =   300
         Visible         =   0   'False
         Width           =   1215
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
      TabIndex        =   0
      Top             =   10620
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2470
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19226
            MinWidth        =   19226
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "2018-04-13"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "오후 1:07"
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
   Begin VB.Frame Frame1 
      Height          =   9165
      Left            =   90
      TabIndex        =   1
      Top             =   1350
      Width           =   15195
      Begin VB.CheckBox chkAll 
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   660
         TabIndex        =   45
         Top             =   300
         Width           =   195
      End
      Begin FPSpread.vaSpread vasID 
         Height          =   8805
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   7935
         _Version        =   393216
         _ExtentX        =   13996
         _ExtentY        =   15531
         _StockProps     =   64
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   17
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmInterface.frx":4301
         UserResize      =   2
      End
      Begin FPSpread.vaSpread vasRes 
         Height          =   8805
         Left            =   8250
         TabIndex        =   47
         Top             =   240
         Width           =   6795
         _Version        =   393216
         _ExtentX        =   11986
         _ExtentY        =   15531
         _StockProps     =   64
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
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
         SpreadDesigner  =   "frmInterface.frx":8624
      End
      Begin VB.Label Label2 
         Caption         =   "Barcode"
         Height          =   255
         Left            =   270
         TabIndex        =   58
         Top             =   270
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.Frame FrmHideControl 
      Caption         =   "HideControl"
      Height          =   4455
      Left            =   540
      TabIndex        =   2
      Top             =   5820
      Visible         =   0   'False
      Width           =   13095
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1575
         Left            =   3540
         TabIndex        =   30
         Top             =   2400
         Width           =   9285
         Begin VB.TextBox txtEquipID 
            Height          =   345
            Left            =   3600
            TabIndex        =   41
            Text            =   "10"
            Top             =   1140
            Width           =   1875
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Rack Pos"
            Height          =   375
            Left            =   7560
            TabIndex        =   40
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command10 
            Caption         =   "결과입력"
            Height          =   375
            Left            =   5880
            TabIndex        =   39
            Top             =   1110
            Width           =   1635
         End
         Begin VB.TextBox txtEquipCode 
            Height          =   345
            Left            =   1710
            TabIndex        =   38
            Text            =   "0ADVI120"
            Top             =   1125
            Width           =   1875
         End
         Begin VB.CommandButton Command9 
            Caption         =   "장비ID조회"
            Height          =   375
            Left            =   60
            TabIndex        =   37
            Top             =   1110
            Width           =   1635
         End
         Begin VB.CommandButton Command8 
            Caption         =   "미검사상세목록"
            Height          =   375
            Left            =   5010
            TabIndex        =   36
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command7 
            Caption         =   "미검사목록"
            Height          =   375
            Left            =   3360
            TabIndex        =   35
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command6 
            Caption         =   "검사상세목록"
            Height          =   375
            Left            =   1710
            TabIndex        =   34
            Top             =   690
            Width           =   1635
         End
         Begin VB.TextBox txtID 
            Height          =   345
            Left            =   6660
            TabIndex        =   33
            Text            =   "05111000003"
            Top             =   720
            Width           =   1875
         End
         Begin VB.CommandButton Command5 
            Caption         =   "검사목록"
            Height          =   375
            Left            =   60
            TabIndex        =   32
            Top             =   690
            Width           =   1635
         End
         Begin VB.CommandButton Command4 
            Caption         =   "서버시간"
            Height          =   375
            Left            =   60
            TabIndex        =   31
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label lblDate2 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   1920
            TabIndex        =   43
            Top             =   330
            Width           =   945
         End
         Begin VB.Label lblDate1 
            AutoSize        =   -1  'True
            Caption         =   "서버시간1"
            Height          =   195
            Left            =   3150
            TabIndex        =   42
            Top             =   330
            Width           =   945
         End
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   210
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   1950
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   285
         Left            =   60
         TabIndex        =   26
         Top             =   555
         Width           =   1635
      End
      Begin VB.TextBox txtTemp 
         Height          =   345
         Left            =   240
         TabIndex        =   25
         Top             =   1380
         Width           =   3045
      End
      Begin VB.Frame Frame3 
         Height          =   585
         Left            =   60
         TabIndex        =   18
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
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
            TabIndex        =   23
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
            TabIndex        =   22
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   1950
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdUp 
         Height          =   525
         Left            =   1260
         Picture         =   "frmInterface.frx":C44F
         Style           =   1  '그래픽
         TabIndex        =   15
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton cmdDown 
         Height          =   525
         Left            =   2010
         Picture         =   "frmInterface.frx":C57E
         Style           =   1  '그래픽
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Command13"
         Height          =   285
         Left            =   1710
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   285
         Left            =   1710
         TabIndex        =   12
         Top             =   570
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox Text3 
         Height          =   345
         Left            =   240
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1500
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   285
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
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
         TabIndex        =   6
         Top             =   930
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox Text2 
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Top             =   2355
         Visible         =   0   'False
         Width           =   3045
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   285
         Left            =   1710
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
      End
      Begin FPSpread.vaSpread vasTemp1 
         Height          =   1125
         Left            =   10740
         TabIndex        =   4
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
         SpreadDesigner  =   "frmInterface.frx":C6B0
      End
      Begin FPSpread.vaSpread vasTemp 
         Height          =   1125
         Left            =   3450
         TabIndex        =   8
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
         SpreadDesigner  =   "frmInterface.frx":10B9E
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1125
         Left            =   7110
         TabIndex        =   11
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
         SpreadDesigner  =   "frmInterface.frx":10E0A
      End
      Begin FPSpread.vaSpread vasList 
         Height          =   1125
         Left            =   5295
         TabIndex        =   24
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
         SpreadDesigner  =   "frmInterface.frx":11076
      End
      Begin FPSpread.vaSpread vasResTemp 
         Height          =   1125
         Left            =   8925
         TabIndex        =   29
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
         SpreadDesigner  =   "frmInterface.frx":112E2
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
         TabIndex        =   44
         Top             =   2370
         Visible         =   0   'False
         Width           =   120
      End
   End
   Begin VB.Menu MnMain 
      Caption         =   "파일"
      Begin VB.Menu MnExit 
         Caption         =   "종료"
      End
      Begin VB.Menu MnTest 
         Caption         =   "테스트"
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

Dim gMT As String
Dim gComState As Long
Dim gErrState As Long


Function LRC(ByVal asData As String) As String
    Dim i As Integer
    Dim a
    
    a = Asc(Left(asData, 1))
    
    For i = 2 To Len(asData)
        a = a Xor Asc(Mid(asData, i, 1))
    Next i
    
    If a = 3 Then a = 127
    
    LRC = Chr(a)
End Function

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
    
    SQL = "select barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, count(*), count(*), max(recedate)" & _
          " from pat_res " & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag in ('B','C') " & vbCrLf & _
          "group by diskno, posno, barcode, seqno, receno, pid, pname, page, psex, jumin, sendflag "
    SQL = SQL & vbCrLf & " Union " & vbCrLf
    SQL = SQL & vbCrLf & _
          "select barcode, seqno, receno, diskno, posno, pid, pname, page, psex, jumin, sendflag, count(*), '0',  max(recedate)" & _
          " from pat_res " & _
          "where examdate = '" & Format(CDate(Text_Today.Text), "yyyymmdd") & "' " & vbCrLf & _
          "  and equipno = '" & Trim(gEquip) & "' " & vbCrLf & _
          "  and sendflag not in ('B','C') " & vbCrLf & _
          "group by diskno, posno, barcode, seqno, receno,  pid, pname, page, psex, jumin, sendflag " & vbCrLf & _
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
        Case "B", "C"
            SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
            SetText vasID, "완료", iRow, colState
'        Case "C"
'            SetBackColor vasID, iRow, iRow, 1, colState, 202, 255, 112
'            SetForeColor vasID, iRow, iRow, colState, colState, 255, 0, 0
'            SetText vasID, "완료(Alarm)", iRow, colState
        Case "O"
            SetText vasID, "오더", iRow, colState
         Case "A"
            SetText vasID, "결과", iRow, colState
        End Select
    Next iRow
End Sub

Private Sub cmdCancelState_Click()
    Dim i As Integer
    Dim lsBarcode As String
    Dim lsDTID As String
    Dim lsStr As String
    
    
    lsDTID = gDoctor(cmbDr.ListIndex).WKPERS_ID
    
    For i = 1 To vasID.DataRowCnt
        vasID.Col = 1
        vasID.Row = i
        If vasID.Value = 1 Then
            lsBarcode = Trim(GetText(vasID, i, colBarCode))
            lsStr = "<NewDataSet>"
            lsStr = lsStr & "    <Table>"
            lsStr = lsStr & "        <QID><![CDATA[PG_SRL.SLP91_U08]]></QID>"
            lsStr = lsStr & "        <QTYPE><![CDATA[Package]]></QTYPE>"
            lsStr = lsStr & "        <USERID><![CDATA[LIA]]></USERID>"
            lsStr = lsStr & "        <EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
            lsStr = lsStr & "        <TABLENAME><![CDATA[]]></TABLENAME>"
            lsStr = lsStr & "        <P0><![CDATA[" & lsBarcode & "]]></P0>"
            lsStr = lsStr & "        <P1><![CDATA[" & lsDTID & "]]></P1>"
            lsStr = lsStr & "        <P2><![CDATA[]]></P2>"
            lsStr = lsStr & "        <P3><![CDATA[]]></P3>"
            lsStr = lsStr & "    </Table>"
            lsStr = lsStr & "</NewDataSet>"
            
            Online_Result_Qry_Conf_Cancel lsStr
            
        End If
    Next
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
    
'    If MSComm1.PortOpen = True Then CX_Init
    
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

Private Sub cmdTest_Click()

    Call ABL500(Mid(Text5.Text, 2))
    Text5.Text = ""
    
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
'    frmUserChange.Show 0
    
End Sub

Private Sub Command15_Click()
    
    Online_XML gXml_S07, "10010700001"
'    vasID.MaxRows = 1
'    SetText vasID, "10010700001", 1, colBarCode
'
'    Get_Sample_Info 1
'
End Sub

Private Sub Command16_Click()
    Dim lsChar As String
    Dim i As Long
    Dim strRece As String
    
'''    strRece = "<NewDataSet><Table>"
'''            strRece = strRece & vbCrLf & "<QID><![CDATA[PG_SRL.SLP91_U06]]></QID>"
'''            strRece = strRece & vbCrLf & "<QTYPE><![CDATA[Package]]></QTYPE>"
'''            strRece = strRece & vbCrLf & "<USERID><![CDATA[LIA]]></USERID>"
'''            strRece = strRece & vbCrLf & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
'''            strRece = strRece & vbCrLf & "<TABLENAME><![CDATA[]]></TABLENAME>"
'''            strRece = strRece & vbCrLf & "<P0><![CDATA[EM014]]></P0>"
'''            strRece = strRece & vbCrLf & "<P1><![CDATA[" & gsPID & "]]></P1>"
'''            strRece = strRece & vbCrLf & "<P2><![CDATA[" & Format(Date, "yyyymmdd") & " " & Format(Time, "hhmmss") & "]]></P2>"
'''            strRece = strRece & vbCrLf & "<P3><![CDATA[SVAN]]></P3>"
'''            strRece = strRece & vbCrLf & "<P4><![CDATA[" & gDoctor(cmbDr.ListIndex).WKPERS_ID & "]]></P4>"
'''            strRece = strRece & vbCrLf & "<P5><![CDATA[X0010]]></P5>"
'''            strRece = strRece & vbCrLf & "<P6><![CDATA[10120]]></P6>"
'''            strRece = strRece & vbCrLf & "<P7><![CDATA[]]></P7>"
'''            strRece = strRece & vbCrLf & "<P8><![CDATA[]]></P8>"
'''            strRece = strRece & vbCrLf & "<P9><![CDATA[]]></P9>"
'''            strRece = strRece & vbCrLf & "<P10><![CDATA[]]></P10>"
'''            strRece = strRece & vbCrLf & "</Table></NewDataSet>"
'''
'''            Online_Param gXml_U06, strRece
'''            gsBarCode = gEMRBarcode
'''
    
    
'    For i = 1 To Len(Text4.Text)
'
'        lsChar = Mid(Text4.Text, i, 1)
'
'        Select Case lsChar
'        Case chrENQ
'            SaveData "[RX]" & lsChar
'            MSComm1.Output = chrACK
'            SaveData "[TX]" & chrACK
'            txtData.Text = ""
'            gRow = -1
'        Case chrSTX
'            txtData.Text = lsChar
'        Case chrLF
'            txtData.Text = txtData.Text & lsChar
'            SaveData "[RX]" & txtData.Text
'            ABL500 Mid(txtData.Text, 2)
'            MSComm1.Output = chrACK
'            SaveData "[TX]" & chrACK
'            txtData.Text = ""
'
'        Case Else
'            txtData.Text = txtData.Text & lsChar
'        End Select
'
'    Next
    ABL500 Text4.Text
    Text4.Text = ""
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

Private Sub Form_Load()
    Dim sDate As String
    Dim i As Integer
    
    If App.PrevInstance Then
        End
    End If
    
    Me.Left = 0
    Me.Top = 0
    
    imgPort.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
    imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon

    cmdReset_Click
    
    GetSetup
    
    SocketsInitialize
    gIP = GetTheIP
    SocketsCleanup
    
    StatusBar1.Panels(1) = gIP
    
    MSComm1.CommPort = gSetup.gPort
'    MSComm1.RTSEnable = gSetup.gRTSEnable
'    MSComm1.DTREnable = gSetup.gDTREnable
    MSComm1.Settings = gSetup.gSpeed & "," & gSetup.gParity & "," & gSetup.gDataBit & "," & gSetup.gStopBit

    If MSComm1.PortOpen = False Then
        MSComm1.PortOpen = True
    End If
    
    If MSComm1.PortOpen Then
        frmInterface.StatusBar1.Panels(2).Text = "COM" & MSComm1.CommPort & " 포트에 연결 되었습니다"
        imgPort.Picture = imlStatus.ListImages("RUN").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("STOP").ExtractIcon
    Else
        frmInterface.StatusBar1.Panels(2).Text = "통신포트에 연결 되지 않았습니다"
        imgPort.Picture = imlStatus.ListImages("STOP").ExtractIcon
        imgSend.Picture = imlStatus.ListImages("NOT").ExtractIcon
        imgReceive.Picture = imlStatus.ListImages("NOT").ExtractIcon
    End If
    
    If Not Connect_Local Then
        MsgBox "연결되지 않았습니다."
        cn_Local_Flag = False
        Exit Sub
    Else
        cn_Local_Flag = True
    End If
    
'    If Not Connect_Local_1 Then
'        MsgBox "연결되지 않았습니다."
'
'    End If
    
        
'    If Not Connect_Server Then
'        MsgBox "연결되지 않았습니다."
'        cn_Server_Flag = False
'        Exit Sub
'    Else
'        cn_Server_Flag = True
'    End If

'    cn_Server_Flag = dce_setenv("client.env", "", "")
'
'    Dim i_equip_cd$
'    Dim machine_id$(), equip_cd$(), equip_nm$()
'
'    'i_equip_cd = gEquip
'    res = sl_sel_machine_id(gEquip, machine_id(), equip_cd(), equip_nm())
'    If res > 0 Then
'        gEquipID = machine_id(0)
'    End If
'    SQL = "update equipexam set equipno = 'C063' where equipno = '" & gEquip & "'"
'    res = SendQuery(gLocal, SQL)

    Text_Today = Format(CDate(GetDateFull), "yyyy/mm/dd")

    GetExamCode
        
    sDate = Format(DateAdd("y", CDate(Text_Today.Text), -365), "yyyymmdd")
    
    SQL = "delete from pat_res where examdate < '" & sDate & "'"
    res = SendQuery(gLocal, SQL)
    
    Online_XML gXml_S26, "SVAN"
    For i = 0 To giIndex
        cmbDr.AddItem gDoctor(i).WKPERS_NM, i
    Next
    
    cmbDr.ListIndex = 0
    
End Sub

Function Get_Sample_Info(ByVal asRow As Long) As Integer
Dim lsBarcode As String
Dim lsPID As String
Dim lsReceNo As String
Dim sRes As String


    Get_Sample_Info = -1
    
    '샘플 환자 정보 가져오기
    
    lsBarcode = Trim(GetText(vasID, asRow, colBarCode))   '샘플 바코드 번호
    
'    If Trim(lsbarcode) = "" Then: Exit Function
    sRes = Online_XML(gXml_S03, lsBarcode)
'    If sRes = 1 Then
        SetText vasID, gPat_Info_Select.PT_NO, asRow, colPID
        SetText vasID, gPat_Info_Select.PT_NM, asRow, colPName
        SetText vasID, gPat_Info_Select.SEX, asRow, colPSex
        SetText vasID, gPat_Info_Select.AGE, asRow, colPAge
        SetText vasID, gPat_Info_Select.ACPTNO_1, asRow, colSeqNo
        SetText vasID, Format(gPat_Info_Select.ACPT_DTETM, "yyyymmdd"), asRow, colDate
        SetText vasID, gPat_Info_Select.SPC_CD_1, asRow, colReceno

'        vasID.RowHeight(asRow) = 20
        
        Get_Sample_Info = 1
'    End If
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
    
    SQL = "Select equipcode, examcode, examname, reflow, refhigh " & vbCrLf & _
          "From equipexam " & vbCrLf & _
          "Where equipno = '" & gEquip & "' " & vbCrLf & _
          "order by  examcode "
    res = db_select_Vas(gLocal, SQL, vasTemp)
    If res > 0 Then
        ReDim gArrEquip(1 To vasTemp.DataRowCnt, 1 To 6)
    Else
        SaveQuery SQL
        Exit Function
    End If
        
    For i = 1 To vasTemp.DataRowCnt
        gArrEquip(i, 1) = i
        For j = 1 To 5
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

Private Sub MnTest_Click()

    If cmdTest.Visible = False Then
        cmdTest.Visible = True
        Text5.Visible = True
    Else
        Text5.Text = ""
        cmdTest.Visible = False
        Text5.Visible = False
    End If
    
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
    
    Select Case MSComm1.CommEvent
        Case comEvReceive

            imgReceive.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrReceive.Enabled = False Then
                tmrReceive.Enabled = True
            Else
                tmrReceive.Enabled = False
                tmrReceive.Enabled = True
            End If
        Case comEvSend
            imgSend.Picture = imlStatus.ListImages("RUN").ExtractIcon
            If tmrSend.Enabled = False Then
                tmrSend.Enabled = True
            Else
                tmrSend.Enabled = False
                tmrSend.Enabled = True
            End If
    End Select
    
    lsChar = MSComm1.Input
    
    StatusBar1.Panels(2) = lsChar
    
    Select Case lsChar
    Case chrENQ
        SaveData "[RX]" & lsChar
        MSComm1.Output = chrACK
        SaveData "[TX]" & chrACK
        txtData.Text = ""
        gRow = -1
    Case chrSTX
        txtData.Text = lsChar
    Case chrLF
        txtData.Text = txtData.Text & lsChar
        SaveData "[RX]" & txtData.Text
        ABL500 Mid(txtData.Text, 2)
        MSComm1.Output = chrACK
        SaveData "[TX]" & chrACK
        txtData.Text = ""
        
    Case Else
        txtData.Text = txtData.Text & lsChar
    End Select
    
End Sub

Sub ABL500(asData As String)
    
    Dim ResultTbl(1 To 40) As String
    Dim TablePtr As Integer
    Dim sTmp As String
    
    Dim i As Integer
    Dim ii As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    
    Dim iCnt As Integer
    
    Dim lsID As String
    Dim lsPID As String
    Dim lsPName As String
    Dim lsJumin1 As String
    Dim lsJumin2 As String
    Dim lsPSex As String
    Dim lsPage As String

    Dim lsTestID As String
    Dim lsExamCode As String
    Dim lsResult As String
    Dim lsExamDate As String
    Dim lsEquipRes As String
    
    Dim sSampleType As String
    Dim sLotNo As String
    Dim sLevel As String
    
    Dim rv As Integer
    Dim vTemp As String
    Dim qOrdDate As String
    Dim qQMCode As String
    Dim qOrdSeqNo As String
    Dim qEquipCode As String
    Dim qSpcCode As String
    Dim qExamCode As String
    Dim qSetYN As String
    Dim qLotNo As String
    Dim qRoomCode As String
    Dim qQCType As String
    Dim qEditID As String
    Dim qEditIP As String
    Dim qTransStr As String
    Dim strRece As String

'〓 구분자
'io_erryn in out varchar2
'io_errmsg in out varchar2
'io_spcidlist in out varchar2
'io_outcnt in out varchar2
'io_rowcnt in out varchar2

    If asData = "" Then
        Exit Sub
    End If

    
    TablePtr = 1
' ----- for start
    For j = 1 To Len(asData)
        If (Mid(asData, j, 1) = "|") Then
            TablePtr = TablePtr + 1
            ResultTbl(TablePtr) = " "
        Else
            ResultTbl(TablePtr) = ResultTbl(TablePtr) + Mid(asData, j, 1)
        End If
    Next j
' ------- for end
    
    If Mid(ResultTbl(1), 2, 1) = "H" Then     'Header Record
        Var_Clear
        
        iCnt = 0
        
        For i = 1 To Len(asData)
            If Mid(asData, i, 1) = "|" Then
                iCnt = iCnt + 1

                Select Case iCnt
                    Case 13
                        gDate = Mid(asData, i + 1, 14)      '장비에서 받은 날짜시간
                End Select
            End If
        Next i
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "O" Then
        sTmp = Trim(ResultTbl(4))      'Sample구분
        i = InStr(1, sTmp, "^")
        If i > 0 Then
            gsSampleType = Mid(sTmp, 1, i - 1)
            If gsSampleType = "Sample #" Then
                gsBarCode = gsSampleType & Mid(sTmp, i + 1)
                gsSampleType = "P"
                
            ElseIf gsSampleType = "QC #" Then
                gsSampleType = "Q"
                sLotNo = Trim(ResultTbl(16)) 'lotno
                i = InStr(1, sLotNo, "")
                If i > 0 Then
                    sLotNo = Mid(sLotNo, 1, i - 1)
                End If
                i = InStr(1, sLotNo, "^")
                If i > 0 Then
'                    sLevel = Mid(sLotNo, 1, i - 1)
'                    sLotNo = Mid(sLotNo, i + 1)
                    sLotNo = Mid(sLotNo, 1, i - 1)
                    
                    sLotNo = Trim(mGetP(mGetP(asData, 16, "|"), 2, "^"))
                    sLotNo = Replace(sLotNo, vbCr, "")
                    sLotNo = Replace(sLotNo, vbLf, "")
                    sLotNo = Replace(sLotNo, chrETX, "")
                    
                    If Len(sLotNo) <> 5 Then
                        sLotNo = Mid(sLotNo, 1, 5)
                    End If
                End If
            Else
                gsSampleType = "C"
            End If
        Else
            gsSampleType = "P"
        End If

        If gsSampleType = "Q" Then  'QC 자동접수
            Online_TLA gXml_S11, gQCEquip, sLotNo
            qOrdDate = ""
            qQMCode = ""
            qOrdSeqNo = ""
            qEquipCode = ""
            qSpcCode = ""
            qExamCode = ""
            qSetYN = ""
            qLotNo = ""
            qRoomCode = ""
            qQCType = ""
            qEditID = ""
            qEditIP = ""
            qQCType = "I"
            qEditID = "10738"
            qEditIP = gIP
            
            For X = 0 To giIndex
                qOrdDate = qOrdDate & gQC_Select(X).ORDDATE & "〓"
                qQMCode = qQMCode & gQC_Select(X).QMCODE & "〓"
                qOrdSeqNo = qOrdSeqNo & gQC_Select(X).ORDSEQNO & "〓"
                qEquipCode = qEquipCode & gQC_Select(X).EQIPCODE & "〓"
                qSpcCode = qSpcCode & gQC_Select(X).SPCCODE & "〓"
                qExamCode = qExamCode & gQC_Select(X).EXAMCODE & "〓"
                qSetYN = qSetYN & gQC_Select(X).SETEXYN & "〓"
                qLotNo = qLotNo & gQC_Select(X).LOTNO & "〓"
                qRoomCode = qRoomCode & gQC_Select(X).ROOMCODE & "〓"
'                If X = 0 Then
'                    qOrdDate = gQC_Select(X).ORDDATE
'                    qQMCode = gQC_Select(X).QMCODE
'                    qOrdSeqNo = gQC_Select(X).ORDSEQNO
'                    qEquipCode = gQC_Select(X).EQIPCODE
'                    qSpcCode = gQC_Select(X).SPCCODE
'                    qExamCode = gQC_Select(X).EXAMCODE
'                    qSetYN = gQC_Select(X).SETEXYN
'                    qLotNo = gQC_Select(X).LOTNO
'                    qRoomCode = gQC_Select(X).ROOMCODE
'
'                Else
'                    qOrdDate = qOrdDate & "〓" & gQC_Select(X).ORDDATE
'                    qQMCode = qQMCode & "〓" & gQC_Select(X).QMCODE
'                    qOrdSeqNo = qOrdSeqNo & "〓" & gQC_Select(X).ORDSEQNO
'                    qEquipCode = qEquipCode & "〓" & gQC_Select(X).EQIPCODE
'                    qSpcCode = qSpcCode & "〓" & gQC_Select(X).SPCCODE
'                    qExamCode = qExamCode & "〓" & gQC_Select(X).EXAMCODE
'                    qSetYN = qSetYN & "〓" & gQC_Select(X).SETEXYN
'                    qLotNo = qLotNo & "〓" & gQC_Select(X).LOTNO
'                    qRoomCode = qRoomCode & "〓" & gQC_Select(X).ROOMCODE
'
'                End If
            Next
            
               
            qTransStr = "<Table>" & vbCrLf & _
                        "<QID><![CDATA[PG_SRL.SLP91_U03]]></QID>" & vbCrLf & _
                        "<QTYPE><![CDATA[Package]]></QTYPE>" & vbCrLf & _
                        "<USERID><![CDATA[LIA]]></USERID>" & vbCrLf & _
                        "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & vbCrLf & _
                        "<TABLENAME><![CDATA[]]></TABLENAME>" & vbCrLf & _
                        "<P0><![CDATA[" & qOrdDate & "]]></P0>" & vbCrLf & _
                        "<P1><![CDATA[" & qQMCode & "]]></P1>" & vbCrLf & _
                        "<P2><![CDATA[" & qOrdSeqNo & "]]></P2>" & vbCrLf & _
                        "<P3><![CDATA[" & qEquipCode & "]]></P3>" & vbCrLf & _
                        "<P4><![CDATA[" & qSpcCode & "]]></P4>" & vbCrLf & _
                        "<P5><![CDATA[" & qExamCode & "]]></P5>" & vbCrLf & _
                        "<P6><![CDATA[" & qSetYN & "]]></P6>" & vbCrLf & _
                        "<P7><![CDATA[" & qLotNo & "]]></P7>" & vbCrLf & _
                        "<P8><![CDATA[" & qRoomCode & "]]></P8>" & vbCrLf & _
                        "<P9><![CDATA[" & qQCType & "]]></P9>" & vbCrLf & _
                        "<P10><![CDATA[" & qEditID & "]]></P10>" & vbCrLf & _
                        "<P11><![CDATA[" & qEditIP & "]]></P11>" & vbCrLf & _
                        "<P12><![CDATA[]]></P12>" & vbCrLf & _
                        "<P13><![CDATA[]]></P13>" & vbCrLf & _
                        "<P14><![CDATA[]]></P14>" & vbCrLf & _
                        "<P15><![CDATA[]]></P15>" & vbCrLf & _
                        "<P16><![CDATA[]]></P16>" & vbCrLf & _
                        "</Table>"
    
            qTransStr = "<NewDataSet>" & qTransStr & "</NewDataSet>"
    
            Online_Param gXml_U03, qTransStr
            
            gsBarCode = gQC_Rece(2)
        
        ElseIf gsSampleType = "C" Then  'Cal
        
        Else
            
            strRece = "<NewDataSet><Table>"
            strRece = strRece & vbCrLf & "<QID><![CDATA[PG_SRL.SLP91_U06]]></QID>"
            strRece = strRece & vbCrLf & "<QTYPE><![CDATA[Package]]></QTYPE>"
            strRece = strRece & vbCrLf & "<USERID><![CDATA[LIA]]></USERID>"
            strRece = strRece & vbCrLf & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
            strRece = strRece & vbCrLf & "<TABLENAME><![CDATA[]]></TABLENAME>"
            strRece = strRece & vbCrLf & "<P0><![CDATA[EM014]]></P0>"
            strRece = strRece & vbCrLf & "<P1><![CDATA[" & gsPID & "]]></P1>"
            strRece = strRece & vbCrLf & "<P2><![CDATA[" & Format(Date, "yyyymmdd") & " " & Format(Time, "hhmmss") & "]]></P2>"
            strRece = strRece & vbCrLf & "<P3><![CDATA[SVAN]]></P3>"
            strRece = strRece & vbCrLf & "<P4><![CDATA[" & gDoctor(cmbDr.ListIndex).WKPERS_ID & "]]></P4>"
            strRece = strRece & vbCrLf & "<P5><![CDATA[X0010]]></P5>"
            strRece = strRece & vbCrLf & "<P6><![CDATA[10120]]></P6>"
            strRece = strRece & vbCrLf & "<P7><![CDATA[]]></P7>"
            strRece = strRece & vbCrLf & "<P8><![CDATA[]]></P8>"
            strRece = strRece & vbCrLf & "<P9><![CDATA[]]></P9>"
            strRece = strRece & vbCrLf & "<P10><![CDATA[]]></P10>"
            strRece = strRece & vbCrLf & "</Table></NewDataSet>"
            
            Online_Param gXml_U06, strRece
            gsBarCode = gEMRBarcode
            
        End If
        
        gOrderExam = ""
        
        gRow = -1
        For i = 1 To vasID.DataRowCnt
'            If sSampleType = "P" Then
                If Trim(GetText(vasID, i, colBarCode)) = gsBarCode Then
                    gRow = i
                    Exit For
                End If
'            ElseIf sSampleType = "Q" Then

'            End If
        Next i
        
        If gRow < 0 Then
            gRow = vasID.DataRowCnt + 1
            If vasID.MaxRows < gRow Then
                vasID.MaxRows = gRow
            End If
        End If
        
        SetText vasID, gsBarCode, gRow, colBarCode
        
        vasActiveCell vasID, gRow, colBarCode
        ClearSpread vasRes
        
        '샘플정보 가져오기
        If gsSampleType = "Q" Then
            SetText vasID, sLotNo, gRow, colPID
            SetText vasID, "QC", gRow, colPName
            
            Online_XML gXml_S07, gsBarCode
        ElseIf gsSampleType = "C" Then
            SetText vasID, "CAL", gRow, colPName
        
        Else
            If Trim(GetText(vasID, gRow, colPID)) = "" And Len(Trim(GetText(vasID, gRow, colBarCode))) = 11 Then
                '검체번호에 대한 검사목록 전체조회 (상세조회)
                Get_Sample_Info gRow

            End If
            Online_XML gXml_S07, gsBarCode
        End If
    End If
    
    
    If (Mid(ResultTbl(1), 2, 1) = "P") Then          'Test Order Record
        sTmp = Trim(ResultTbl(4))      '검체번호
        i = InStr(1, sTmp, Chr(13))
        gsPID = ""
        If i > 0 Then
            gsPID = Mid(sTmp, 1, i - 1)
        Else
            gsPID = sTmp
        End If
    End If
    
    If Mid(ResultTbl(1), 2, 1) = "L" Then
        gOrderExam = ""
        If MnTransAuto.Checked = True Then
        
            res = Insert_Data(gRow)
            
            If res = -1 Then
                SetForeColor vasID, gRow, gRow, 1, colState, 255, 0, 0
                SetText vasID, "실패", gRow, colState
            Else
               
                SetBackColor vasID, gRow, gRow, 1, colState, 202, 255, 112
                SetText vasID, "완료", gRow, colState
                
                SQL = " Update pat_res Set " & vbCrLf & _
                      " sendflag = 'C' " & vbCrLf & _
                      " Where equipno = '" & gEquip & "' " & vbCrLf & _
                      " And barcode = '" & Trim(GetText(vasID, gRow, colBarCode)) & "' "
                res = SendQuery(gLocal, SQL)
                If res = -1 Then
                    SaveQuery SQL
                    Exit Sub
                End If
                
            End If
                    
        End If
    
    End If
    

    If (Mid(ResultTbl(1), 2, 1) = "R") Then     'Result
        gOrderMessage = "R"
        
        sTmp = ResultTbl(3)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        sTmp = Mid(sTmp, i + 1)
        i = InStr(1, sTmp, "^")
        lsTestID = Left(sTmp, i - 1)    '장비코드

        sTmp = ResultTbl(4)
        lsResult = Trim(sTmp)           '결과
        
        
        gsResDateTime = ResultTbl(10)    'result time
        
        ClearSpread vasTemp
        
        SQL = "Select examcode, examname From equipexam" & vbCrLf & _
              "Where equipno = '" & gEquip & "' " & vbCrLf & _
              "And equipcode = '" & lsTestID & "'  "
        If Mid(gsBarCode, 5, 1) <> "9" Then
            SQL = SQL & "and examcode like 'X%' "
        Else
            SQL = SQL & "and examcode like 'L%' "
        End If
        
        res = db_select_Vas(gLocal, SQL, vasTemp)
            
        If vasTemp.DataRowCnt > 0 Then
            k = -1
            If vasTemp.DataRowCnt > 1 And vasResTemp.DataRowCnt > 0 Then
                For j = 1 To vasTemp.DataRowCnt
                    k = -1
                    For X = 1 To vasResTemp.DataRowCnt
                        If Trim(GetText(vasResTemp, X, 1)) = Trim(GetText(vasTemp, j, 1)) Then
                            k = j
                            Exit For
                        End If
                    Next X
                    If k > 0 Then
                        Exit For
                    End If
                Next j
            End If
            
            If k < 1 Then
                k = 1
            Else
                vasTemp.MaxRows = k
            End If
            
            For j = k To vasTemp.DataRowCnt
'                i = i + 1
                
'                If IsNumeric(lsTestID) = True And IsNumeric(lsResult) = True Then
                    i = vasRes.DataRowCnt + 1
                    
                    If i > vasRes.MaxRows Then
                        vasRes.MaxRows = i
                    End If
                    
                    If i > 0 Then
                        lsExamCode = Trim(GetText(vasTemp, j, 1))
                        
                        '숫자만 디스플레이 하기
                        If IsNumeric(lsResult) = False Then
                            For ii = 1 To Len(lsResult)
                                If Mid(lsResult, ii, 1) = "?" Then
                                    lsResult = Mid(lsResult, ii + 1)
                                    
                                    Exit For
                                End If
                            Next ii
                        End If
                        
                        '소수점 처리, 결과 형태 처리
                        
                        lsEquipRes = lsResult
                        lsResult = SetResult(lsResult, lsTestID)
                        
                        SetText vasRes, lsTestID, i, colEquipCode                           '장비코드
                        SetText vasRes, Trim(GetText(vasTemp, j, 1)), i, colExamCode        '검사코드
                        SetText vasRes, Trim(GetText(vasTemp, j, 2)), i, colExamName        '검사명
                        
                        SetText vasRes, lsResult, i, colResult                              '변환결과
                        SetText vasRes, lsEquipRes, i, colResult1                           '장비결과
                        
                        Save_Local_One_1 gRow, i, "A"
                    End If
               
'                End If
            Next j
        End If
        
    End If
    
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
    
    If IsNumeric(sEquipRes) = False Then
        Exit Function
    End If
    
    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' "
    res = db_select_Col(gLocal, SQL)
    
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
    
    If IsNumeric(gReadBuf(1)) = True Then
        sLVal = gReadBuf(1)
        If CCur(sLVal) > CCur(sEquipRes) Then
            sResFlag = "<"
        End If
    End If
    
    If IsNumeric(gReadBuf(2)) = True Then
        sHVal = gReadBuf(2)
        If CCur(sHVal) < CCur(sEquipRes) Then
            sResFlag = ">"
        End If
    End If
    
    sResult = sResFlag & sResult
    SetResult = sResult
    
End Function

Function Save_Local_One_1(ByVal asRow1 As Long, ByVal asRow2 As Long, asSend As String)
    Dim sCnt As String
    Dim sExamDate As String

    sExamDate = GetDateFull
    
    If UCase(Left(Trim(GetText(vasID, asRow1, colPJumin)), 1)) = "F" Then
'        Save_Local_QC Trim(Text_Today.Text) & " " & Format(Time, "hh:nn:ss"), _
                      Trim(GetText(vasID, asRow1, colBarcode)), _
                      Trim(GetText(vasRes, asRow2, colEquipCode)), _
                      Trim(GetText(vasRes, asRow2, colResult)), _
                      Trim(GetText(vasRes, asRow2, colResult1))
        'Save_Local_QC Trim(Text_Today.Text) & " " & Trim(GetText(vasID, asRow1, colPID)), _
                      Trim(GetText(vasID, asRow1, colBarcode)), _
                      Trim(GetText(vasRes, asRow2, colEquipCode)), _
                      Trim(GetText(vasRes, asRow2, colResult)), _
                      Trim(GetText(vasRes, asRow2, colResult1))
        Exit Function
    End If

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
    Dim sRes As String
    
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
    
    gOrderExam = ""
    
'    Online_XML gXml_S07, lsID

    SQL = "Select equipcode, examcode, examname, result, result1 " & vbCrLf & _
          "from pat_res " & vbCrLf & _
          "where equipno = '" & gEquip & "' " & vbCrLf & _
          "  and barcode = '" & lsID & "' " & vbCrLf & _
          "  and result <> '' "
'    SQL = "Select equipcode, examcode, examname, result, result1 " & vbCrLf & _
'          "from pat_res " & vbCrLf & _
'          "where equipno = '" & gEquip & "' " & vbCrLf & _
'          "  and barcode = '" & lsID & "' " & vbCrLf & _
'          "  and examcode in (" & gOrderExam & ") " & vbCrLf & _
'          "  and result <> '' "
    If asSend = 0 Then
'        SQL = SQL & vbCrLf & _
'          "  and sendflag <> 'C' "
    End If
    
    res = db_select_Vas(gLocal, SQL, vasTemp)
    
    If vasTemp.DataRowCnt < 1 Then Exit Function

    Save_Raw_Data lsID & " : 서버 결과 전송 시작"
    Save_Raw_Data lsID & " : 장부 정보 가져오기"

    On Error GoTo ErrHandle
    
    sParam = ""
    
    For sResRow = 1 To vasTemp.DataRowCnt
        If Trim(GetText(vasTemp, sResRow, 2)) <> "" And IsNumeric(Trim(GetText(vasTemp, sResRow, 4))) = True Then
            sParam = sParam & "<Table>" & _
                    "<QID><![CDATA[PG_SRL.SLP91_P03]]></QID>" & _
                    "<QTYPE><![CDATA[Package]]></QTYPE>" & _
                    "<USERID><![CDATA[LIA]]></USERID>" & _
                    "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>" & _
                    "<TABLENAME><![CDATA[]]></TABLENAME>" & _
                    "<P0><![CDATA[" & lsID & "]]></P0>" & _
                    "<P1><![CDATA[" & Trim(GetText(vasTemp, sResRow, 2)) & "]]></P1>" & _
                    "<P2><![CDATA[" & Trim(GetText(vasTemp, sResRow, 4)) & "]]></P2>" & _
                    "<P3><![CDATA[]]></P3>" & _
                    "<P4><![CDATA[" & gEquip & "]]></P4>" & _
                    "<P5><![CDATA[]]></P5>" & _
                    "<P6><![CDATA[]]></P6>" & _
                    "<P7><![CDATA[]]></P7>" & _
                    "<P8><![CDATA[]]></P8>" & _
                    "<P9><![CDATA[]]></P9>" & _
                    "</Table>"
'            SQL = "Update pat_res set sendflag = 'C' " & vbCrLf & _
'                  "where equipno = '" & gEquip & "' " & vbCrLf & _
'                  "  and barcode = '" & lsID & "' and examcode = '" & Trim(GetText(vasTemp, sResRow, 2)) & "'"
'            res = SendQuery(gLocal, SQL)
'            If res = -1 Then
'                SaveQuery SQL
'                Exit Function
'            End If
        End If
    Next
    
    sParam = "<NewDataSet>" & sParam & "</NewDataSet>"
    
    Online_Result_Qry sParam
    
    Insert_Data = 1
    
    sParam = "<NewDataSet><Table>"
    sParam = sParam & vbCrLf & "<QID><![CDATA[PG_SRL.SLP91_U07]]></QID>"
    sParam = sParam & vbCrLf & "<QTYPE><![CDATA[Package]]></QTYPE>"
    sParam = sParam & vbCrLf & "<USERID><![CDATA[LIA]]></USERID>"
    sParam = sParam & vbCrLf & "<EXECTYPE><![CDATA[FILL]]></EXECTYPE>"
    sParam = sParam & vbCrLf & "<TABLENAME><![CDATA[]]></TABLENAME>"
    sParam = sParam & vbCrLf & "<P0><![CDATA[" & lsID & "]]></P0>"
    sParam = sParam & vbCrLf & "<P1><![CDATA[10120]]></P1>"
    sParam = sParam & vbCrLf & "<P2><![CDATA[]]></P2>"
    sParam = sParam & vbCrLf & "<P3><![CDATA[]]></P3></Table></NewDataSet>"
    
    Online_Result_Qry_Conf sParam
    
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
'    If KeyCode = vbKeyReturn Then
'        LX20 Text2
'        'MsgBox CS(Text2)
'
'        'Hitachi747 Mid(Text2.Text, 2)
'        'txtData = ""
'    End If
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
    Dim lsReceNo As String
    Dim lsSeqno As String
    
    
    Dim iRow As Long
    
    'cmdCall_Click
    
    If Row < 1 Or Row > vasID.DataRowCnt Then
        Exit Sub
    End If
    
    
    lsID = Trim(GetText(vasID, Row, colBarCode))
    lsReceNo = Trim(GetText(vasID, Row, colReceno))
    lsSeqno = Trim(GetText(vasID, Row, colSeqNo))
    
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
          "  and a.barcode = '" & lsID & "' and a.receno = '" & lsReceNo & "'  and a.seqno = '" & lsSeqno & "' " & vbCrLf & _
          "  and a.examcode <> a.equipcode " & vbCrLf & _
          "  and b.equipno = a.equipno " & vbCrLf & _
          "  and b.equipcode = a.equipcode " & vbCrLf & _
          "  and b.examcode = a.examcode"
    res = db_select_Vas(gLocal, SQL, vasRes)
    SQL = "Select a.equipcode, a.examcode, max(b.examname), a.result, b.seqno, a.refflag, a.result1 " & vbCrLf & _
          "from pat_res a, equipexam b " & vbCrLf & _
          "where a.examdate = '" & Format(CDate(Text_Today), "yyyymmdd") & "' " & vbCrLf & _
          "  and a.equipno = '" & gEquip & "' " & vbCrLf & _
          "  and a.barcode = '" & lsID & "' and a.receno = '" & lsReceNo & "'  and a.seqno = '" & lsSeqno & "' " & vbCrLf & _
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



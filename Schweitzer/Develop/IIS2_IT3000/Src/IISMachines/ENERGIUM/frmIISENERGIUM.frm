VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIISENERGIUM 
   BackColor       =   &H00FFFFFF&
   Caption         =   "INFINITY"
   ClientHeight    =   9180
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   15735
   ScaleWidth      =   28680
   StartUpPosition =   3  'Windows 기본값
   Begin MSWinsockLib.Winsock wSck 
      Left            =   6660
      Top             =   8460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPtid 
      Alignment       =   2  '가운데 맞춤
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8550
      TabIndex        =   48
      Text            =   "00576711"
      Top             =   90
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CheckBox chkDualTest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dual Test"
      Height          =   255
      Left            =   11400
      TabIndex        =   47
      Top             =   180
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CheckBox chkAll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "전체접수"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2610
      TabIndex        =   46
      Top             =   465
      Width           =   1050
   End
   Begin VB.CheckBox chkTimer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Refresh"
      Height          =   255
      Left            =   12690
      TabIndex        =   43
      Top             =   180
      Value           =   1  '확인
      Width           =   960
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00DBE6E6&
      Caption         =   "전송"
      Height          =   315
      Left            =   3960
      Style           =   1  '그래픽
      TabIndex        =   42
      Top             =   90
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdResult 
      BackColor       =   &H00FFFFFF&
      Caption         =   "받기"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9540
      Style           =   1  '그래픽
      TabIndex        =   41
      Top             =   465
      Width           =   900
   End
   Begin VB.Timer tmrOrder 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7785
      Top             =   8460
   End
   Begin VB.CommandButton cmdConfig 
      BackColor       =   &H00DBE6E6&
      Caption         =   "설 정(&S)"
      Height          =   405
      Left            =   8865
      Style           =   1  '그래픽
      TabIndex        =   40
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtResultSec 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13095
      MaxLength       =   9
      TabIndex        =   38
      Text            =   "60"
      Top             =   495
      Width           =   615
   End
   Begin VB.TextBox txtOrderSec 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5850
      MaxLength       =   9
      TabIndex        =   36
      Text            =   "60"
      Top             =   495
      Width           =   615
   End
   Begin VB.Frame fraHidden 
      Caption         =   "숨김"
      Height          =   5055
      Left            =   7470
      TabIndex        =   10
      Top             =   2925
      Visible         =   0   'False
      Width           =   4830
      Begin VB.CommandButton cmdMake 
         Caption         =   "Temp생성"
         Height          =   285
         Left            =   2925
         TabIndex        =   45
         Top             =   675
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CheckBox chkPOC 
         BackColor       =   &H00DBE6E6&
         Caption         =   "POCT"
         Height          =   255
         Left            =   2925
         TabIndex        =   44
         Top             =   315
         Value           =   1  '확인
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00DBE6E6&
         Height          =   1290
         Left            =   135
         TabIndex        =   18
         Top             =   1785
         Width           =   8595
         Begin MedControls1.LisLabel lblPtId 
            Height          =   315
            Left            =   1245
            TabIndex        =   19
            Top             =   165
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   "00000001"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblDoctNm 
            Height          =   315
            Left            =   3930
            TabIndex        =   20
            Top             =   165
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   "이상대"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblStatFg 
            Height          =   315
            Left            =   6795
            TabIndex        =   21
            Top             =   165
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   "응급"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblName 
            Height          =   315
            Left            =   1245
            TabIndex        =   22
            Top             =   525
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   "이상대 아기"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblDeptNm 
            Height          =   315
            Left            =   3930
            TabIndex        =   23
            Top             =   525
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   "수술실"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblSpcNm 
            Height          =   315
            Left            =   6795
            TabIndex        =   24
            Top             =   525
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   "Blood"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblSexAge 
            Height          =   315
            Left            =   1245
            TabIndex        =   25
            Top             =   885
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   "남자 / 29"
            Appearance      =   0
         End
         Begin MedControls1.LisLabel lblWardNm 
            Height          =   315
            Left            =   3930
            TabIndex        =   26
            Top             =   885
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   0
            Alignment       =   1
            Caption         =   "65병동"
            Appearance      =   0
         End
         Begin VB.Label lblControl 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "환  자 ID :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   150
            TabIndex        =   34
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lblLevel 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "이     름 :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   150
            TabIndex        =   33
            Top             =   600
            Width           =   990
         End
         Begin VB.Label lblLotNo 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "성별/나이 :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   150
            TabIndex        =   32
            Top             =   975
            Width           =   990
         End
         Begin VB.Label lblGeneral 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "처방의 :"
            BeginProperty Font 
               Name            =   "굴림체"
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
            TabIndex        =   31
            Top             =   240
            Width           =   720
         End
         Begin VB.Label lblGeneral 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "진료과 :"
            BeginProperty Font 
               Name            =   "굴림체"
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
            TabIndex        =   30
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lblGeneral 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "병  동 : "
            BeginProperty Font 
               Name            =   "굴림체"
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
            TabIndex        =   29
            Top             =   975
            Width           =   810
         End
         Begin VB.Label lblGeneral 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "응급여부 :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   5760
            TabIndex        =   28
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblGeneral 
            AutoSize        =   -1  'True
            BackColor       =   &H00DBE6E6&
            Caption         =   "검 체 명 :"
            BeginProperty Font 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   4
            Left            =   5760
            TabIndex        =   27
            Top             =   600
            Width           =   900
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   345
         Left            =   1485
         TabIndex        =   14
         Top             =   945
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.FileListBox FileENERGIUM 
         Height          =   270
         Left            =   135
         Pattern         =   "*.dat"
         TabIndex        =   13
         Top             =   1005
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtWorkNo 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1935
         MaxLength       =   9
         TabIndex        =   11
         Top             =   450
         Width           =   615
      End
      Begin MedControls1.LisLabel LisLabel1 
         Height          =   375
         Left            =   135
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1485
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   661
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "■ 환자정보"
         Appearance      =   0
      End
      Begin FPSpread.vaSpread tblResult 
         Height          =   6690
         Left            =   180
         TabIndex        =   35
         Top             =   3240
         Width           =   8580
         _Version        =   393216
         _ExtentX        =   15134
         _ExtentY        =   11800
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   22
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         ScrollBars      =   2
         SelectBlockOptions=   0
         ShadowColor     =   13697023
         SpreadDesigner  =   "frmIISENERGIUM.frx":0000
         TextTip         =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Start WorkNo : "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   405
         TabIndex        =   12
         Top             =   525
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "조회"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3735
      Style           =   1  '그래픽
      TabIndex        =   9
      Top             =   465
      Width           =   900
   End
   Begin VB.Timer tmrResult 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   8235
      Top             =   8460
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   405
      Left            =   11310
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      Height          =   405
      Left            =   10095
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox txtBarNo 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      TabIndex        =   0
      Text            =   "123456789011"
      Top             =   510
      Width           =   1530
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   405
      Left            =   12525
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8520
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   8880
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MDIControls.MDIActiveX MDIActiveX 
      Left            =   7238
      Top             =   8507
      _ExtentX        =   847
      _ExtentY        =   794
   End
   Begin FPSpread.vaSpread tblReady 
      Height          =   8040
      Left            =   105
      TabIndex        =   4
      Top             =   855
      Width           =   6480
      _Version        =   393216
      _ExtentX        =   11430
      _ExtentY        =   14182
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   6
      MaxRows         =   10
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmIISENERGIUM.frx":06D7
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   7425
      Left            =   6615
      TabIndex        =   5
      Top             =   855
      Width           =   7110
      _Version        =   393216
      _ExtentX        =   12541
      _ExtentY        =   13097
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   14
      MaxRows         =   14
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   16777215
      SpreadDesigner  =   "frmIISENERGIUM.frx":0C55
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   98
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   107
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "■ 검사대상 리스트"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Left            =   6630
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   105
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      Caption         =   "■ 검사완료 리스트"
      Appearance      =   0
   End
   Begin MSComCtl2.DTPicker dtpFromDt 
      Height          =   330
      Left            =   1185
      TabIndex        =   15
      Top             =   495
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      Format          =   132120577
      CurrentDate     =   38330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Refresh : "
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12105
      TabIndex        =   39
      Top             =   555
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Refresh : "
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4860
      TabIndex        =   37
      Top             =   555
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "▶ 접수일자"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   16
      Top             =   510
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "바코드번호 : "
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6675
      TabIndex        =   8
      Top             =   525
      Width           =   1065
   End
End
Attribute VB_Name = "frmIISENERGIUM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISENERGIUM.frm
'   작성자  : 오세원
'   내  용  : ENERGIUM 장비폼
'   작성일  : 2021-08-12
'   버  전  :
'   병  원  :
'       1. 전주예수병원
'   메  모  : NOTE 확인할것!
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady의 Column Enum
Private Enum TReadyEnum
    ccNo = 1
    ccBarNo = 2
    ccAccNo = 3
    ccPtId = 4
    ccName = 5
End Enum

'## tblComplete의 Column Enum
Private Enum TCompleteEnum
    ccNo = 1:           ccBarNo = 2
    ccAccNo = 3:        ccPtId = 4
    ccName = 5:         ccSexAge = 6
    ccDoctNm = 7:       ccDeptNm = 8
    ccWardNm = 9:       ccStatFg = 10
    ccSpcNm = 11:       ccQcFg = 12
    ccSendCnt = 13:     ccResult = 14
End Enum

'## tblResult의 Column Enum
Private Enum TResultEnum
    ccTestNm = 1
    ccEqpResult = 2
    ccLISResult = 3
    ccUnit = 4
    ccHLDiv = 5
    ccDPDiv = 6
    ccRef = 7
    ccInfo = 8
End Enum

'## Clear Enum
Private Enum ClearEnum
    ccAll = 1
    ccLabel = 2
End Enum

'## Popup Menu ID
Private Const DELETE    As Long = 1
Private Const DELETEALL As Long = 2

'## Datalog Field 상수
Private Const Rs As String = ""    'Record Separator
Private Const FS As String = ""    'Field Separator
Private Const GS As String = ""    'Group Separator

Private WithEvents mIntLib  As clsIISInterface   '인터페이스 클래스
Attribute mIntLib.VB_VarHelpID = -1
Private WithEvents mPopup   As clsIISPopup       '팝업메뉴
Attribute mPopup.VB_VarHelpID = -1

Private mIntErrors  As clsIISIntErrors           '인터페이스 에러 컬렉션
Private mOrder      As clsIISIntOrder            '오더정보 클래스

Private mEqpCd  As String   '장비코드
Private mEqpKey As String   '장비키

Private lngOrder As Long
Private lngResult As Long
Private blnRS   As Boolean

Private AdoCn           As ADODB.Connection
'Private DBCon As clsIISDbCon

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
     ByVal CodePage As Long, _
     ByVal dwFlags As Long, _
     ByVal lpWideCharStr As Long, _
     ByVal cchWideChar As Long, _
     ByVal lpMultiByteStr As Long, _
     ByVal cchMultiByte As Long, _
     ByVal lpDefaultChar As Long, _
     ByVal lpUsedDefaultChar As Long _
) As Long

 

Private Const CP_UTF8 As Long = 65001

Public Function URLEncodeUTF8(Str As String) As String
On Error GoTo ErrLbl

     Dim BufSize As Long, MultiArr() As Byte, Buf As String, i As Long
     Dim UniArr() As Byte
     UniArr = Str
    
     BufSize = WideCharToMultiByte(CP_UTF8, 0&, VarPtr(UniArr(0)), (UBound(UniArr) + 1) / 2, 0&, 0&, 0&, 0&)
    
     If BufSize > 0 Then
          ReDim MultiArr(BufSize - 1&)
          WideCharToMultiByte CP_UTF8, 0&, VarPtr(UniArr(0)), (UBound(UniArr) + 1) / 2, VarPtr(MultiArr(0)), BufSize, 0&, 0&
     End If
    
'for문엔에 MultiArr(i)의 문자코드 값이 알파벳인지를
'구별하셔야 알파벳 이외의 문자들만 변환하셔야.
'변환전에 전체를 변환하지 마시고 알파벳 이외의 문자들이
'나타날때만 변환하는 코드로 적용하심이
     
     For i = 0 To UBound(MultiArr)
        Debug.Print MultiArr(i)
        If MultiArr(i) > 127 Then
            Buf = Buf & "%" & Hex$(MultiArr(i))
        Else
            Buf = Buf & MultiArr(i)
        End If
     Next i
    
     URLEncodeUTF8 = Buf

ErrLbl:
End Function


Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property


Private Sub chkTimer_Click()

    If chkTimer.Value = 1 Then
        tmrOrder.Enabled = True
        tmrResult.Enabled = True
    Else
        tmrOrder.Enabled = False
        tmrResult.Enabled = False
    End If
    
End Sub

Private Sub cmdConfig_Click()
    With frmIISConfig
        .EqpKey = mEqpKey
        .Show vbModal
    End With
End Sub

Private Sub cmdMake_Click()
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim pBarNo As String
    Dim intRow As Integer
    
    Screen.MousePointer = 11
    
    tmrOrder.Enabled = False
    tmrResult.Enabled = False

    With tblReady
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = 2
            pBarNo = Trim(.Text)
            If pBarNo <> "" Then
                'If chkDualTest.Value = "1" Then
                '    Set objAccInfo = mIntLib.GetAccInfo_New_Temp_Dual(pBarNo, chkPOC.Value)
                'Else
                    Set objAccInfo = mIntLib.GetAccInfo_New_Temp(pBarNo, chkPOC.Value)
                'End If
            End If
        Next
    End With
    
    If chkTimer.Value = "1" Then
        tmrOrder.Enabled = True
        tmrResult.Enabled = True
    End If
    
    Screen.MousePointer = 0

End Sub

Private Sub cmdOrder_Click()
    Dim objAccInfo  As clsIISAccInfo    '접수내역 객체
    Dim objPro      As clsIISProgress   'Progressbar
    Dim vOrdChk     As Variant          'Spread의 전송유무
    Dim vPID        As Variant          'Spread의 환자번호
    Dim vBarNo      As Variant          'Spread의 바코드번호
    Dim vAccNo      As Variant          'Spread의
    Dim vPatNm      As Variant
    Dim vSTAT       As Variant
    Dim strFileNm   As String           'JobList 파일명(경로포함)
    Dim strBackUpNm   As String           'JobList 파일명(경로포함)
    Dim strHeader   As String
    Dim strBody     As String
    Dim strRecord   As String           'JobList 파일의 Record
    Dim lngFileNo   As Long             'File No
    Dim i           As Long
    Dim j           As Integer
    Dim k           As Integer
    Dim varRecord   As Variant
    Dim STM As ADODB.Stream
    Dim strAccDt    As String
    Dim strAccseq   As String
    Dim strWA       As String
    Dim strTestCd   As String
    Dim strSql      As String
    Dim strSPC      As String
    Dim strBackUpPath As String
    
    'Set STM = New ADODB.Stream
        
    Me.MousePointer = vbHourglass
    
    '## 오더파일 생성
    With tblReady
        If .DataRowCnt = 0 Then Exit Sub
        
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccNo, i, vOrdChk)
            
            If IsNumeric(vOrdChk) Then
                Call .GetText(TReadyEnum.ccName, i, vPatNm)
                Call .GetText(TReadyEnum.ccPtId, i, vPID)
                Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
                Call .GetText(TReadyEnum.ccAccNo, i, vAccNo)
                Call .GetText(6, i, vSTAT)
                
                strRecord = Get_TestList(CStr(vAccNo))
                If strRecord <> "" Then
                    '^^^LCR132^\^^^LCR131^\^^^LCR130^\^^^LCR123^\^^^LCR121^\^^^LCR126^\^^^LCR124^\^^^LCR144^\^^^LCR125^|
                    strRecord = mGetP(strRecord, 1, "|")
                    strTestCd = mGetP(strRecord, 2, "|")
                End If
                
                '-- 등록된 검사항목일 경우
                If strRecord <> "" Then
                    strFileNm = mOrderPath & "\" & vBarNo & ".dat"
                    strBackUpPath = mBackUpPath & "\" & Format(Now, "yyyymmdd")
                    strBackUpNm = strBackUpPath & "\" & vBarNo & ".dat"
                    
                    If Dir$(strFileNm, vbNormal) <> "" Then
                        Kill strFileNm
                    End If
                    '## 파일오픈
                    Set STM = New ADODB.Stream
                    STM.Open
                    STM.Type = adTypeText
                    STM.Charset = "utf-8"
                    
                    '일반(입원)
                    If CStr(vSTAT) = "입원" Then
                        strRecord = "^^^IN^\" & strRecord
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & strRecord & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf
                        
                        STM.WriteText strHeader
                        
                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        If Dir(strBackUpPath, vbDirectory) <> Format(Now, "yyyymmdd") Then
                            Call MkDir(strBackUpPath)
                        End If
                        FileCopy strFileNm, strBackUpNm
                        
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                        strWA = mGetP(vAccNo, 1, "-")
                        strAccDt = mGetP(vAccNo, 2, "-")
                        strAccseq = mGetP(vAccNo, 3, "-")
                        If strTestCd <> "" Then
                            strSql = ""
                            strSql = strSql & "UPDATE S2LAB320                      " & vbCrLf
                            strSql = strSql & "   SET STSCD     = '1'               " & vbCrLf
                            strSql = strSql & " WHERE WORKAREA  = '" & strWA & "'   " & vbCrLf
                            strSql = strSql & "   AND ACCDT     = '" & strAccDt & "'" & vbCrLf
                            strSql = strSql & "   AND ACCSEQ    = " & strAccseq & vbCrLf
                            strSql = strSql & "   AND TESTCD    IN (" & strTestCd & ") "
                            'strSql = strSql & "   AND TESTCD IN ('B2721','B2602','B2580','B2611','C2462','B2570','C3720','C3721','F6932A','C2261','C2262','C2411','C3793','B2630','C2243','B2710','C2420','C2490','B2590','C2430','B2621','C3797','C2302','C22001','C2443','C2500','C463','C2510','C3780','C4602A','C4602B','C4602','C4903','C2210','F6932E','F6932B','F6932C','F6932D','C3730','C3795','C3750','C3711','C37111G','C37111','C37112','C37112G','C37113G','C3792','C3791','C3794','C3812','C2200') & vbcrlf"

                            AdoCn.Execute strSql, , adCmdText + adExecuteNoRecords
                        End If
                    ElseIf CStr(vSTAT) = "외래" Then
                        strRecord = "^^^OUT^\" & strRecord
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & strRecord & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf
                        
                        STM.WriteText strHeader
                        
                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        If Dir(strBackUpPath, vbDirectory) <> Format(Now, "yyyymmdd") Then
                            Call MkDir(strBackUpPath)
                        End If
                        FileCopy strFileNm, strBackUpNm
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                        strWA = mGetP(vAccNo, 1, "-")
                        strAccDt = mGetP(vAccNo, 2, "-")
                        strAccseq = mGetP(vAccNo, 3, "-")
                        If strTestCd <> "" Then
                            strSql = ""
                            strSql = strSql & "UPDATE S2LAB320                      " & vbCrLf
                            strSql = strSql & "   SET STSCD     = '1'               " & vbCrLf
                            strSql = strSql & " WHERE WORKAREA  = '" & strWA & "'   " & vbCrLf
                            strSql = strSql & "   AND ACCDT     = '" & strAccDt & "'" & vbCrLf
                            strSql = strSql & "   AND ACCSEQ    = " & strAccseq & vbCrLf
                            strSql = strSql & "   AND TESTCD    IN (" & strTestCd & ") "
                            AdoCn.Execute strSql, , adCmdText + adExecuteNoRecords
                        End If
                    ElseIf CStr(vSTAT) = "외주" Then
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & "^^^SEND^" & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf
                        
                        STM.WriteText strHeader
                        
                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        If Dir(strBackUpPath, vbDirectory) <> Format(Now, "yyyymmdd") Then
                            Call MkDir(strBackUpPath)
                        End If
                        FileCopy strFileNm, strBackUpNm
                        
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                        strWA = mGetP(vAccNo, 1, "-")
                        strAccDt = mGetP(vAccNo, 2, "-")
                        strAccseq = mGetP(vAccNo, 3, "-")
                        If strTestCd <> "" Then
                            strSql = ""
                            strSql = strSql & "UPDATE S2LAB320                      " & vbCrLf
                            strSql = strSql & "   SET STSCD     = '1'               " & vbCrLf
                            strSql = strSql & " WHERE WORKAREA  = '" & strWA & "'   " & vbCrLf
                            strSql = strSql & "   AND ACCDT     = '" & strAccDt & "'" & vbCrLf
                            strSql = strSql & "   AND ACCSEQ    = " & strAccseq & vbCrLf
                            strSql = strSql & "   AND TESTCD    IN (" & strTestCd & ") "
                            AdoCn.Execute strSql, , adCmdText + adExecuteNoRecords
                        End If
                    ElseIf CStr(vSTAT) = "응급" Then
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & "^^^STAT^" & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf
                        
                        STM.WriteText strHeader
                        
                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        If Dir(strBackUpPath, vbDirectory) <> Format(Now, "yyyymmdd") Then
                            Call MkDir(strBackUpPath)
                        End If
                        FileCopy strFileNm, strBackUpNm
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                        strWA = mGetP(vAccNo, 1, "-")
                        strAccDt = mGetP(vAccNo, 2, "-")
                        strAccseq = mGetP(vAccNo, 3, "-")
                        If strTestCd <> "" Then
                            strSql = ""
                            strSql = strSql & "UPDATE S2LAB320                      " & vbCrLf
                            strSql = strSql & "   SET STSCD     = '1'               " & vbCrLf
                            strSql = strSql & " WHERE WORKAREA  = '" & strWA & "'   " & vbCrLf
                            strSql = strSql & "   AND ACCDT     = '" & strAccDt & "'" & vbCrLf
                            strSql = strSql & "   AND ACCSEQ    = " & strAccseq & vbCrLf
                            strSql = strSql & "   AND TESTCD    IN (" & strTestCd & ") "
                            AdoCn.Execute strSql, , adCmdText + adExecuteNoRecords
                        End If
                    ElseIf CStr(vSTAT) = "ARCH2" Then
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & "^^^ARCH2^" & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf
                        
                        STM.WriteText strHeader
                        
                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        If Dir(strBackUpPath, vbDirectory) <> Format(Now, "yyyymmdd") Then
                            Call MkDir(strBackUpPath)
                        End If
                        FileCopy strFileNm, strBackUpNm
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                        strWA = mGetP(vAccNo, 1, "-")
                        strAccDt = mGetP(vAccNo, 2, "-")
                        strAccseq = mGetP(vAccNo, 3, "-")
                        If strTestCd <> "" Then
                            strSql = ""
                            strSql = strSql & "UPDATE S2LAB320                      " & vbCrLf
                            strSql = strSql & "   SET STSCD     = '1'               " & vbCrLf
                            strSql = strSql & " WHERE WORKAREA  = '" & strWA & "'   " & vbCrLf
                            strSql = strSql & "   AND ACCDT     = '" & strAccDt & "'" & vbCrLf
                            strSql = strSql & "   AND ACCSEQ    = " & strAccseq & vbCrLf
                            strSql = strSql & "   AND TESTCD    IN (" & strTestCd & ") "
                            AdoCn.Execute strSql, , adCmdText + adExecuteNoRecords
                        End If
                    End If
                Else
                    strFileNm = mOrderPath & "\" & vBarNo & ".dat"
                    strBackUpPath = mBackUpPath & "\" & Format(Now, "yyyymmdd")
                    strBackUpNm = strBackUpPath & "\" & vBarNo & ".dat"
                                        
                    If Dir$(strFileNm, vbNormal) <> "" Then
                        Kill strFileNm
                    End If
                    '## 파일오픈
                    Set STM = New ADODB.Stream
                    STM.Open
                    STM.Type = adTypeText
                    STM.Charset = "utf-8"
                                            
                    '일반
                    If CStr(vSTAT) = "입원" Then
                        strRecord = "^^^IN^\" & strRecord
                        
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & strRecord & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf

                        STM.WriteText strHeader

                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                    ElseIf CStr(vSTAT) = "외래" Then
                        strRecord = "^^^OUT^\" & strRecord
                        
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & strRecord & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf

                        STM.WriteText strHeader

                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                    
                    ElseIf CStr(vSTAT) = "외주" Then
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & "^^^SEND^" & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf
                        
                        STM.WriteText strHeader
                        
                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        If Dir(strBackUpPath, vbDirectory) <> Format(Now, "yyyymmdd") Then
                            Call MkDir(strBackUpPath)
                        End If
                        FileCopy strFileNm, strBackUpNm
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                    ElseIf CStr(vSTAT) = "응급" Then
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & "^^^STAT^" & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf
                        
                        STM.WriteText strHeader
                        
                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        If Dir(strBackUpPath, vbDirectory) <> Format(Now, "yyyymmdd") Then
                            Call MkDir(strBackUpPath)
                        End If
                        FileCopy strFileNm, strBackUpNm
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                    ElseIf CStr(vSTAT) = "ARCH2" Then
                        strHeader = ""
                        strHeader = strHeader & "H|\^&|||ASTM-Host|||||PSM||P||" & Format(Now, "yyyymmddhhmmss") & vbCrLf
                        strHeader = strHeader & "P|1||" & vPID & "||" & vPatNm & "^|||||||||||||||||||||||||||||" & vbCrLf
                        strHeader = strHeader & "O|1|" & vBarNo & "||" & "^^^ARCH2^" & "||" & Format(Now, "yyyymmddhhmmss") & "|||||A||||||||||||||O|||||" & vbCrLf
                        strHeader = strHeader & "L|1|F" & vbCrLf
                        
                        STM.WriteText strHeader
                        
                        Call mIntLib.WriteLog("[전송]" & strHeader & strBody, ccEqp)
                        STM.SaveToFile strFileNm, adSaveCreateNotExist
                        STM.Close
                        Set STM = Nothing
                        If Dir(strBackUpPath, vbDirectory) <> Format(Now, "yyyymmdd") Then
                            Call MkDir(strBackUpPath)
                        End If
                        FileCopy strFileNm, strBackUpNm
                        Call .SetText(TReadyEnum.ccNo, i, "전송")
                    End If
                End If
            End If
        Next i
    End With
    
    Set objPro = Nothing
    
    Me.MousePointer = vbDefault
    
    'MsgBox "정상적으로 " & mOrderFileNm & " 파일이 생성되었습니다.", vbInformation, "정보"
End Sub

Private Sub cmdResult_Click()

    Dim intRow      As Integer
    Dim intIdx      As Integer
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim strBuffer   As String
    Dim strtmpBuf   As String
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim intCnt      As Integer
    Dim varTmp      As Variant

On Error GoTo ErrRoutine

    tmrOrder.Enabled = False
    tmrResult.Enabled = False

    FileENERGIUM.Refresh

    DoEvents

    For intIdx = 0 To FileENERGIUM.ListCount - 1
        FileENERGIUM.ListIndex = intIdx
        strSrcfile = FileENERGIUM.Path & FileENERGIUM.FileName     ' 원본 파일 이름을 정의합니다.
        Open strSrcfile For Input As #3

        strBuffer = ""
        Do While Not EOF(3)
            strBuffer = strBuffer & Input(1, #3)
        Loop

        Close #3

        '대상 파일 이름을 정의
        strDestFile = App.Path & "\Log\" & FileENERGIUM.FileName
        '원본을 대상에 복사
        FileCopy strSrcfile, strDestFile

        Kill strSrcfile
        FileENERGIUM.Refresh

        mIntLib.Phase = 1

        Call mIntLib.WriteLog(strBuffer, ccEqp)
        lngBufLen = Len(strBuffer)

        If strBuffer <> "" Then
            Call mIntLib.AddBuffer(strBuffer)
            Call EditRcvData
        End If
    Next

    If chkTimer.Value = "1" Then
        tmrOrder.Enabled = True
        tmrResult.Enabled = True
    End If
    
Exit Sub

ErrRoutine:
    
    If chkTimer.Value = "1" Then
        tmrOrder.Enabled = True
        tmrResult.Enabled = True
    End If
    
End Sub

Private Sub cmdSearch_Click()

    If chkAll.Value = "1" Then
        If MsgBox("검체량에 따라 30분이상 걸릴수 있습니다." & vbNewLine & dtpFromDt.Value & "일 전체접수를 진행하시겠습니까?", vbYesNo + vbInformation + vbDefaultButton2, Me.Caption) = vbYes Then
            tmrOrder.Enabled = False
            tmrResult.Enabled = False
            
            '   Spread Initializing
            Call Set_SpreadInit(tblReady)
            
            Call Get_SearchList
        
            Call cmdMake_Click
            
            Call cmdOrder_Click
            
            If chkTimer.Value = "1" Then
                tmrOrder.Enabled = True
                tmrResult.Enabled = True
            End If
            
            chkAll.Value = "0"
        
        End If
    Else
        tmrOrder.Enabled = False
        tmrResult.Enabled = False
        
        '   Spread Initializing
        Call Set_SpreadInit(tblReady)
        
        Call Get_SearchList
    
        Call cmdMake_Click
        
        Call cmdOrder_Click
        
        If chkTimer.Value = "1" Then
            tmrOrder.Enabled = True
            tmrResult.Enabled = True
        End If
    End If
End Sub


'   Access DB Connect
Public Function Set_DbConnect_DB() As Boolean
    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean
    Dim strSrcfile  As String
    Dim strDestFile As String
    Dim strCon As String
    
    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError

    DB_Name = mDB
    UserName = mUID
    Password = mPW

    If (DB_Name = "") Or (UserName = "") Then
        Set_DbConnect_DB = False
        Set AdoCn = Nothing
        Exit Function
    End If
        
    strCon = "Provider=MSDAORA.1;Persist Security Info=True;" & _
         "Data Source=" & DB_Name & ";" & _
         "User ID=" & UserName & ";" & _
         "Password=" & Password
        
    With AdoCn
        .ConnectionString = ""
'        .ConnectionTimeout = 25
        .CursorLocation = adUseClient
'        .Provider = "Microsoft.Jet.OLEDB.4.0"
'        .Properties("Mode").Value = adModeReadWrite
'        .Properties("Persist Security Info").Value = False
'        .Properties("Data Source").Value = DB_Name
'        .Properties("User ID").Value = UserName
'        .Properties("Jet OLEDB:Database Password").Value = Password
'        .Properties("Jet OLEDB:Compact Without Replica Repair").Value = True
        .Open strCon
    
    
    End With

    Set_DbConnect_DB = True
    
 Exit Function

ConnectError:
    '   오류처리
    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf & vbCrLf _
           , vbCritical, " DB Open Error"

    If AdoCn.State <> adStateOpen Then
        Set_DbConnect_DB = False
        Set AdoCn = Nothing
    End If

End Function

'   Record Set Open
Public Function Get_Recordset(ByVal AdoCn As ADODB.Connection, ByVal strSql As String, _
                             ByVal AdoRS As ADODB.Recordset, _
                             Optional Call_Name As String, _
                             Optional Cursor_Location As ADODB.CursorLocationEnum = adUseClient, _
                             Optional Cursor_Type As ADODB.CursorTypeEnum = adOpenStatic, _
                             Optional Lock_Type As ADODB.LockTypeEnum = adLockPessimistic) As Boolean

On Error GoTo DBOpenRsError
    
    With AdoRS
        .CursorLocation = Cursor_Location
        .Source = strSql
        .ActiveConnection = AdoCn
        .CursorType = Cursor_Type
        .LockType = Lock_Type
        .Open
    End With
    
    Get_Recordset = True

Exit Function

DBOpenRsError:
    Set AdoRS = Nothing
    Get_Recordset = False

End Function

'   Newest Result Recordset
Public Function Get_NewResult() As ADODB.Recordset
    Dim strSql      As String
    Dim AdoRS       As ADODB.Recordset
    Dim strPTID     As String
    
    strPTID = txtPtid.Text
    
    strSql = ""
    strSql = strSql & "SELECT b.PATNAME, a.*                " & vbCrLf
    strSql = strSql & "  FROM S2LAB201 a,  ORAA1.APPATBAT b " & vbCrLf
    strSql = strSql & " WHERE a.ACCDT   = '" & Format(dtpFromDt.Value, "yyyymmdd") & "'" & vbCrLf
    strSql = strSql & "   AND a.PTID    = b.PATNO           " & vbCrLf
    '테스트용
'    If strPTID <> "" Then
'        strSql = strSql & "   AND a.PTID    = '" & strPTID & "' " & vbCrLf
'        strSql = strSql & "   AND ROWNUM    = 1                 " & vbCrLf
'    End If
    
    strSql = strSql & "   AND a.SPCCD IN ('1A','1H','3A','3I','3C','4L','4E','1J','4B','S1','S2','S3','S4','S5','S6','S7','S8','S9','30','40','4P','4V','4T','T3','U1','3D','1X') " & vbCrLf
    If chkAll.Value = 0 Then
        '??????? 채혈상태도 가져온다....?????
        strSql = strSql & "   AND a.STSCD <= '2'    " & vbCrLf
        'strSql = strSql & "   AND a.STSCD  = '2'    " & vbCrLf
        strSql = strSql & "   AND NOT EXISTS        " & vbCrLf
        strSql = strSql & "       (SELECT WORKAREA || ACCDT || ACCSEQ   " & vbCrLf
        strSql = strSql & "          From S2LAB320                      " & vbCrLf
        strSql = strSql & "         Where WORKAREA  = A.WORKAREA        " & vbCrLf
        strSql = strSql & "           AND ACCDT     = A.ACCDT           " & vbCrLf
        strSql = strSql & "           AND ACCSEQ    = A.ACCSEQ          " & vbCrLf
        strSql = strSql & "           AND ACCDT     = '" & Format(dtpFromDt.Value, "yyyymmdd") & "')" & vbCrLf
    End If
    strSql = strSql & " ORDER BY a.SPCYY, a.SPCNO, b.PATNAME, a.ACCDT, a.ACCSEQ  " & vbCrLf

'    Call mIntLib.WriteLog("[대상자조회]" & vbCrLf & strSql, ccEqp)

    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(DbCon, strSql, AdoRS, "") Then
        Set Get_NewResult = AdoRS
        blnRS = True
    Else
        Set Get_NewResult = Nothing
        blnRS = False
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing
    blnRS = False

End Function


'   Newest Result Recordset
Public Function Get_NewResult_Barcode(ByVal pBarNo As String) As ADODB.Recordset
    Dim strSql      As String
    Dim AdoRS       As ADODB.Recordset
    Dim strSpcYy    As String
    Dim lngSpcNo    As Long
    
    strSpcYy = Mid$(pBarNo, 1, SPCYYLEN)
    lngSpcNo = CLng(Mid$(pBarNo, SPCYYLEN + 1, SPCNOLEN))
    
             strSql = "SELECT b.PATNAME, a.* " & vbCr
    strSql = strSql & "  FROM S2LAB201 a,  ORAA1.APPATBAT b" & vbCr
'    strSql = strSql & " WHERE a.ACCDT = '" & Format(dtpFromDt.Value, "yyyymmdd") & "'" & vbCr
'    strSql = strSql & "   AND a.STSCD <= '2' " & vbCr
    strSql = strSql & " WHERE a.SPCYY = '" & strSpcYy & "'"
    strSql = strSql & "   AND a.SPCNO = " & lngSpcNo
    strSql = strSql & "   AND a.PTID = b.PATNO " & vbCr
    strSql = strSql & "   AND a.SPCCD IN ('1A','1H','3A','3I','3C','4L','4E','1J','4B','S1','S2','S3','S4','S5','S6','S7','S8','S9','30','40','4P','4V','4T','T3','U1','3D','1X') "
'    strSql = strSql & "   AND NOT EXISTS"
'    strSql = strSql & "       (SELECT WORKAREA || ACCDT || ACCSEQ"
'    strSql = strSql & "          From S2LAB320"
'    strSql = strSql & "         Where WORKAREA = A.WORKAREA"
'    strSql = strSql & "           AND ACCDT = A.ACCDT"
'    strSql = strSql & "           AND ACCSEQ = A.ACCSEQ"
'    strSql = strSql & "           AND ACCDT = '" & Format(dtpFromDt.Value, "yyyymmdd") & "')"
    strSql = strSql & " ORDER BY a.SPCYY, a.SPCNO, b.PATNAME, a.ACCDT, a.ACCSEQ  "

    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(DbCon, strSql, AdoRS, "") Then
        Set Get_NewResult_Barcode = AdoRS
        blnRS = True
    Else
        Set Get_NewResult_Barcode = Nothing
        blnRS = False
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing
    blnRS = False

End Function

Public Function Get_SendFlag(ByVal pAccNo As String) As Boolean
    Dim strSql      As String
    Dim AdoRS       As ADODB.Recordset
    Dim strAccDt    As String
    Dim strAccseq   As String
    Dim strWA       As String
    
    Get_SendFlag = False
    
    strWA = mGetP(pAccNo, 1, "-")
    strAccDt = mGetP(pAccNo, 2, "-")
    strAccseq = mGetP(pAccNo, 3, "-")
             
             strSql = "SELECT STSCD         " & vbCrLf
    strSql = strSql & "  FROM S2LAB320      " & vbCrLf
    strSql = strSql & " WHERE WORKAREA  = '" & strWA & "'" & vbCrLf
    strSql = strSql & "   AND ACCDT     = '" & strAccDt & "'" & vbCrLf
    strSql = strSql & "   AND ACCSEQ    = " & strAccseq & vbCrLf
    strSql = strSql & "   AND TESTCD IN (SELECT INTBASE FROM S2LAB702 WHERE EQPCD = '" & mEqpCd & "')" & vbCrLf
    strSql = strSql & "   AND SPCCD IN ('1A','1H','3A','3I','3C','4L','4E','1J','4B','S1','S2','S3','S4','S5','S6','S7','S8','S9','30','40','4P','4V','4T','T3','U1','3D','1X') " & vbCrLf
    strSql = strSql & " ORDER BY TESTCD  " & vbCrLf
    
    'Call mIntLib.WriteLog("[대상자별 상태조회]" & vbCrLf & strSql, ccEqp)

    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(DbCon, strSql, AdoRS, "") Then
        If Not AdoRS.BOF Then
            Do Until AdoRS.EOF
                If AdoRS.Fields("STSCD") & "" = "1" Then
                    Get_SendFlag = True
                    Exit Do
                Else
                    Get_SendFlag = False
                End If
                AdoRS.MoveNext
            Loop
        End If
    Else
        Get_SendFlag = False
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Get_SendFlag = False
    Set AdoRS = Nothing

End Function


Public Function Get_TestDetail(ByVal pAccNo As String) As ADODB.Recordset
    Dim strSql      As String
    Dim AdoRS       As ADODB.Recordset
'    Dim strSpcYy    As String
'    Dim lngSpcNo    As Long
    Dim strAccDt    As String
    Dim strAccseq   As String
    Dim strWA       As String
    
    strWA = mGetP(pAccNo, 1, "-")
    strAccDt = mGetP(pAccNo, 2, "-")
    strAccseq = mGetP(pAccNo, 3, "-")
             
             strSql = "SELECT DISTINCT SPCCD, TESTCD        " & vbCrLf
    strSql = strSql & "  FROM S2LAB320                      " & vbCrLf
    strSql = strSql & " WHERE WORKAREA  = '" & strWA & "'   " & vbCrLf
    strSql = strSql & "   AND ACCDT     = '" & strAccDt & "'" & vbCrLf
    strSql = strSql & "   AND ACCSEQ    = " & strAccseq & vbCrLf
    strSql = strSql & "   AND TESTCD IN (SELECT INTBASE FROM S2LAB702 WHERE EQPCD = '" & mEqpCd & "')" & vbCrLf
    strSql = strSql & "   AND SPCCD IN ('1A','1H','3A','3I','3C','4L','4E','1J','4B','S1','S2','S3','S4','S5','S6','S7','S8','S9','30','40','4P','4V','4T','T3','U1','3D','1X') " & vbCrLf
    strSql = strSql & " ORDER BY TESTCD  " & vbCrLf

    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(DbCon, strSql, AdoRS, "") Then
        Set Get_TestDetail = AdoRS
        blnRS = True
    Else
        Set Get_TestDetail = Nothing
        blnRS = False
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing
    blnRS = False

End Function

'   Result List Recordset
'Public Function Get_ResultList() As ADODB.Recordset
'    Dim strSql      As String
'
'On Error GoTo ErrorTrap
'             strSql = "SELECT a.*, b.intnm, b.testcd, b.result, b.hldiv, b.dpdiv "
'    strSql = strSql & "  FROM ACC203 a, ACC204 b "
'    strSql = strSql & " WHERE a.ITEMSEQ = b.ITEMSEQ "
'    strSql = strSql & "   AND a.TRANSDT BETWEEN '" & Format(dtpFrDt.Value, "yyyymmdd") & "' AND '" & Format(dtpToDt.Value, "yyyymmdd") & "'"
'    If Trim(cboSpcPos.Text) <> "ALL" Then
'        strSql = strSql & "   AND a.SPCPOS = '" & cboSpcPos.Text & "'"
'    End If
'    strSql = strSql & " ORDER BY a.ITEMSEQ desc "
'
'    Set AdoRS = New ADODB.Recordset
'    If Get_Recordset(AdoCn, strSql, AdoRS, "") Then
'        Set Get_ResultList = AdoRS
'        blnRS = True
'    Else
'        Set Get_ResultList = Nothing
'        blnRS = False
'    End If
'
'    Set AdoRS = Nothing
'
'Exit Function
'
'ErrorTrap:
'    Set AdoRS = Nothing
'    blnRS = False
'
'End Function

Private Sub Get_SearchList()
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    Dim AdoRS       As ADODB.Recordset
    Dim blnSend     As Boolean
    Dim strAccNo    As String
    
    '   Newest Result
    Set AdoRS = Get_NewResult
    
    If blnRS = False Then
        MsgBox "ENERGIUM 조회 자료가 없습니다.", vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Not AdoRS.BOF Then
        AdoRS.MoveFirst
        tblReady.MaxRows = AdoRS.RecordCount
        intRow = 1: strTransDt = ""
        Do Until AdoRS.EOF
            With tblReady
                '-- 전송한 자료인지 체크(S2LAB320)
                '   Newest Result
                strAccNo = mGetAccNo(AdoRS.Fields("WORKAREA").Value, AdoRS.Fields("ACCDT").Value, AdoRS.Fields("ACCSEQ").Value)
                
                blnSend = Get_SendFlag(strAccNo)
                
                If blnSend Then
                    .SetText 1, intRow, "전송"
                Else
                    .SetText 1, intRow, intRow
                End If
                
                If Trim(AdoRS.Fields("SPCYY").Value) = "" Then
                    .SetText 2, intRow, "-1"
                Else
                    .SetText 2, intRow, Trim(AdoRS.Fields("SPCYY").Value) & Format$(Trim(AdoRS.Fields("SPCNO").Value), String$(SPCNOLEN, "0"))
                End If
                .SetText 3, intRow, strAccNo
                .SetText 4, intRow, AdoRS.Fields("PTID").Value
                .SetText 5, intRow, AdoRS.Fields("PATNAME").Value
                If Trim(AdoRS.Fields("WORKAREA").Value) = "OT" Then
                    .SetText 6, intRow, "외주"
                Else
                    If Trim(AdoRS.Fields("DEPTCD").Value) = "EM" Then
                        .SetText 6, intRow, "응급"
                    Else
                        If Mid(Trim(AdoRS.Fields("SPCCD").Value), 1, 1) = "S" Or Trim(AdoRS.Fields("SPCCD").Value) = "1A" Or Trim(AdoRS.Fields("SPCCD").Value) = "1H" Or Trim(AdoRS.Fields("SPCCD").Value) = "1X" Then
                            .SetText 6, intRow, "일반"
                            If Trim(AdoRS.Fields("ROOMID").Value) <> "" And Not IsNull(Trim(AdoRS.Fields("ROOMID").Value)) Then
                                .SetText 6, intRow, "입원"
                            Else
                                .SetText 6, intRow, "외래"
                            End If
                        Else
                            .SetText 6, intRow, "ARCH2"
                        End If
                    End If
                End If
                intRow = intRow + 1
                AdoRS.MoveNext
            End With
        Loop
    End If
    
    Set AdoRS = Nothing
    
End Sub

Function Get_TestList(ByVal pAccNo As String) As String
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    Dim AdoRS       As ADODB.Recordset
    Dim i As Integer
    Dim strData(2)  As String
    
    Set AdoRS = Get_TestDetail(pAccNo)
    
    If blnRS = False Then
        Exit Function
    End If
    
    strData(1) = ""
    strData(2) = ""
    Get_TestList = ""
    
    If Not AdoRS.BOF Then
        AdoRS.MoveFirst
        Do Until AdoRS.EOF
            '^^^LCR132^\^^^LCR131^\^^^LCR130^\^^^LCR123^\^^^LCR121^\^^^LCR126^\^^^LCR124^\^^^LCR144^\^^^LCR125^
            strData(1) = strData(1) & "\^^^" & Trim(AdoRS.Fields("TESTCD").Value) & "^"
            strData(2) = strData(2) & ",'" & Trim(AdoRS.Fields("TESTCD").Value) & "'"
            
            AdoRS.MoveNext
        Loop
    End If
    
    If strData(1) <> "" Then
        strData(1) = Mid(strData(1), 2)
    End If
    
    If strData(2) <> "" Then
        strData(2) = Mid(strData(2), 2)
    End If
    
    If strData(1) <> "" Or strData(2) <> "" Then
        Get_TestList = strData(1) & "|" & strData(2)
    End If
    
    Set AdoRS = Nothing
    
End Function

'   Spread Initializing
Private Sub Set_SpreadInit(ByVal ClrSpread As Object)
    
    With ClrSpread
        .MaxRows = 0
        .Col = 1
        .Col2 = .MaxCols
        .Row = 1
        .Row2 = .MaxRows
        .BlockMode = True
        .Action = 12    '## ActionClearText
        .BlockMode = False
    End With

End Sub

Private Sub Command1_Click()
    Call MSComm_OnComm
End Sub

''Private Sub Command2_Click()
''    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
''    Dim pBarNo As String
''    Dim intRow As Integer
''
''    Screen.MousePointer = 11
''
''    tmrOrder.Enabled = False
''    tmrResult.Enabled = False
''
''    With tblReady
''        For intRow = 1 To .DataRowCnt
''            .Row = intRow
''            .Col = 2
''            pBarNo = Trim(.Text)
''            If pBarNo <> "" Then
''                Set objAccInfo = mIntLib.GetAccInfo_New_Temp(pBarNo, chkPOC.Value)
''            End If
''        Next
''    End With
''
''    tmrOrder.Enabled = True
''    tmrResult.Enabled = True
''
''    Screen.MousePointer = 0
''
''End Sub

Private Sub Form_Activate()
    MainFrm.lblMenuNm = Me.Caption
    Me.MDIActiveX.WindowState = ccMaximize
End Sub

Private Sub Form_Load()
    On Error GoTo Err
    
    Me.Caption = mEqpKey
    Me.MousePointer = vbHourglass
    
    Set mIntErrors = New clsIISIntErrors
    Set mIntLib = New clsIISInterface
    Set mOrder = New clsIISIntOrder
    
    Call CtlClear
    Call mIntLib.SetConfig(mEqpCd, mEqpKey)
'    Call GetEqpComm
    
    dtpFromDt.Value = Now
    
    txtOrderSec.Text = mOrderRefresh
    txtResultSec.Text = mResultRefresh
    
    lngOrder = mOrderRefresh
    lngResult = mResultRefresh
    
'    If Right(mResultPath, 1) <> "\" Then
'         mResultPath = mResultPath & "\"
'    End If
    
    FileENERGIUM.Path = mResultPath '& "\"
    
    tmrOrder.Interval = 1000    'mOrderRefresh
    tmrResult.Interval = 1000   'mResultRefresh
    
    tmrOrder.Enabled = True
    tmrResult.Enabled = True
    
    DoEvents
    
    If Not Set_DbConnect_DB Then
        'End
    End If
    
    Call Set_SpreadInit(tblReady)
    Call Set_SpreadInit(tblComplete)
        
    chkDualTest.Value = "0"
    
'    If EMPID = "9999" Then
'        txtPtid.Visible = True
'    Else
'        txtPtid.Visible = False
'        txtPtid.Text = ""
'    End If
    
    Me.MousePointer = vbDefault
        
    Exit Sub
Err:
    If Err.Number = "68" Then
        MsgBox mResultPath & " 는 " & Err.Description, vbOKOnly + vbCritical, "ENERGIUM"
        Me.MousePointer = vbDefault
        On Error Resume Next
    Else
    
    End If
End Sub

Private Sub Form_Deactivate()
    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISENERGIUM = Nothing
    Set mOrder = Nothing
End Sub

Private Sub cmdAlarm_Click()
    Dim objErrorShow As clsIISErrorShow     '에러폼 표시 클래스
    
    Set objErrorShow = New clsIISErrorShow
    Call objErrorShow.ShowErrors(mIntErrors)
    Set objErrorShow = Nothing
    
    '## 에러가 없으면 버튼색깔 원래대로, 있으면 계속 빨강색
    cmdAlarm.BackColor = IIf(mIntErrors.Count = 0, &HF4F0F2, vbRed)
    
    '## Alarm창이 닫힌후 포커스를 txtBarNo로 이동
    txtBarNo.SetFocus
    
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
    Call mIntLib.AccInfos.RemoveAll
    
    txtBarNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    tmrResult.Enabled = False
    Unload Me
End Sub

Private Sub tblReady_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    With tblReady
        .Row = Row
        .Col = 1
        If Trim(.Text) <> "" Then
            If Trim(.Text) = "전송" Then
                Call .SetText(1, Row, Row)
            Else
                Call .SetText(1, Row, "전송")
            End If
        End If
    End With

End Sub

Private Sub tmrOrder_Timer()
     
    lngOrder = lngOrder - 1
    If lngOrder = 0 Then
        Call cmdSearch_Click
        
        Call cmdOrder_Click
        
        lngOrder = mOrderRefresh
        txtOrderSec.Text = mOrderRefresh
    Else
        txtOrderSec.Text = lngOrder
    End If
    
    dtpFromDt.Value = Now
End Sub

Private Sub tmrResult_Timer()
    
    lngResult = lngResult - 1
    If lngResult = 0 Then
        Call cmdResult_Click
        lngResult = mResultRefresh
        txtResultSec.Text = mResultRefresh
    Else
        txtResultSec.Text = lngResult
    End If
    
End Sub

Private Sub txtWorkNo_GotFocus()
    With txtWorkNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWorkNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtWorkNo_KeyPress(KeyAscii As Integer)
    '## 숫자만 입력
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txtBarNo_GotFocus()
    With txtBarNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtBarNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    Dim AdoRS       As ADODB.Recordset
    Dim blnSend     As Boolean
    Dim strAccNo    As String
    
    If KeyCode = vbKeyReturn Then
        
        '   Newest Result
        Set AdoRS = Get_NewResult_Barcode(Trim(txtBarNo.Text))
        
        If blnRS = False Then
            MsgBox Trim(txtBarNo.Text) & "의 조회 자료가 없습니다.", vbOKOnly + vbInformation, Me.Caption
            Exit Sub
        End If
        
        If Not AdoRS.BOF Then
            AdoRS.MoveFirst
            strTransDt = ""
            Do Until AdoRS.EOF
                With tblReady
                    '-- 전송한 자료인지 체크(S2LAB320)
                    '   Newest Result
                    .MaxRows = .MaxRows + 1
                    intRow = .MaxRows
                    strAccNo = mGetAccNo(AdoRS.Fields("WORKAREA").Value, AdoRS.Fields("ACCDT").Value, AdoRS.Fields("ACCSEQ").Value)
                    
                    blnSend = Get_SendFlag(strAccNo)
                    
                    If blnSend Then
                        .SetText 1, intRow, "전송"
                    Else
                        .SetText 1, intRow, intRow
                    End If
                    
                    If Trim(AdoRS.Fields("SPCYY").Value) = "" Then
                        .SetText 2, intRow, "-1"
                    Else
                        .SetText 2, intRow, Trim(AdoRS.Fields("SPCYY").Value) & Format$(Trim(AdoRS.Fields("SPCNO").Value), String$(SPCNOLEN, "0"))
                    End If
                    .SetText 3, intRow, strAccNo
                    .SetText 4, intRow, AdoRS.Fields("PTID").Value
                    .SetText 5, intRow, AdoRS.Fields("PATNAME").Value
                    If Trim(AdoRS.Fields("WORKAREA").Value) = "OT" Then
                        .SetText 6, intRow, "외주"
                    Else
                        If Trim(AdoRS.Fields("DEPTCD").Value) = "EM" Then
                            .SetText 6, intRow, "응급"
                        Else
                            If Mid(Trim(AdoRS.Fields("SPCCD").Value), 1, 1) = "S" Or Trim(AdoRS.Fields("SPCCD").Value) = "1A" Or Trim(AdoRS.Fields("SPCCD").Value) = "1H" Or Trim(AdoRS.Fields("SPCCD").Value) = "1X" Then
                                .SetText 6, intRow, "일반"
                            Else
                                .SetText 6, intRow, "ARCH2"
                            End If
                        End If
                    End If
                    'intRow = intRow + 1
                    Call cmdMake_Click
    
                    Call cmdOrder_Click

                    AdoRS.MoveNext
                End With
            Loop
        End If
        
        Set AdoRS = Nothing
        
        txtBarNo.Text = ""
    End If


End Sub

Private Sub txtBarNo_KeyPress(KeyAscii As Integer)
    '## 숫자만 입력
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub tblReady_Click(ByVal Col As Long, ByVal Row As Long)
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim vBarNo      As Variant          'Spread의 바코드번호
    Dim strSpcYy    As String           '검체연도
    Dim lngSpcNo    As Long             '검체번호
    
    If Row = 0 Then Exit Sub
    
    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    Call tblReady.GetText(TReadyEnum.ccBarNo, Row, vBarNo)
    If vBarNo = "" Then Exit Sub
    
    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
'    Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
    
    '## tblResult, Label에 정보표시
'    Call SetLabel(objAccInfo)
'    Call SetResult(objAccInfo)
    
    Set objAccInfo = Nothing
End Sub

Private Sub tblReady_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    Set mPopup = New clsIISPopup
    With mPopup
        .AddMenu DELETE, "Delete"
        .AddMenu DELETEALL, "Delete All"
        .PopupMenus Me.hWnd
    End With
    Set mPopup = Nothing
End Sub

Private Sub tblComplete_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strQcFg     As String   'QC유무
    Dim strIntInfo  As String   '수신한 추가정보
    Dim strTemp     As String
    Dim i           As Long
    
    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    With tblComplete
        .Row = Row
        
        .Col = TCompleteEnum.ccQcFg:    strQcFg = .Text
        If strQcFg = "0" Then
            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
            .Col = TCompleteEnum.ccDoctNm:  lblDoctNm.Caption = .Text
            .Col = TCompleteEnum.ccDeptNm:  lblDeptNm.Caption = .Text
            .Col = TCompleteEnum.ccWardNm:  lblWardNm.Caption = .Text
            .Col = TCompleteEnum.ccStatFg:  lblStatFg.Caption = .Text
            .Col = TCompleteEnum.ccSpcNm:   lblSpcNm.Caption = .Text
        ElseIf strQcFg = "1" Then
            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
        End If
        
        For i = TCompleteEnum.ccResult To .DataColCnt
            .Col = i:   strTemp = .Text
            
            '## 화면표시 버그수정
            If tblResult.MaxRows <= tblResult.DataRowCnt Then
                tblResult.MaxRows = tblResult.MaxRows + 1
                tblResult.Row = tblResult.MaxRows
            Else
                tblResult.Row = tblResult.DataRowCnt + 1
            End If
            
            tblResult.Col = TResultEnum.ccTestNm:       tblResult.Text = mGetP(strTemp, TResultEnum.ccTestNm, DIV)
            tblResult.Col = TResultEnum.ccUnit:         tblResult.Text = mGetP(strTemp, TResultEnum.ccUnit, DIV)
            tblResult.Col = TResultEnum.ccHLDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccHLDiv, DIV)
            tblResult.Col = TResultEnum.ccDPDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccDPDiv, DIV)
            tblResult.Col = TResultEnum.ccRef:          tblResult.Text = mGetP(strTemp, TResultEnum.ccRef, DIV)
            tblResult.Col = TResultEnum.ccInfo:
                strIntInfo = mGetP(strTemp, TResultEnum.ccInfo, DIV)
                tblResult.Text = strIntInfo
            tblResult.Col = TResultEnum.ccEqpResult:
                '## 에러정보가 있는경우 장비결과를 빨강색으로
                tblResult.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
                If Trim(strIntInfo) = "" Then
                    tblResult.ForeColor = vbBlack
                Else
                    tblResult.ForeColor = vbRed
                End If
            tblResult.Col = TResultEnum.ccLISResult:
                '## 결과등록 할수없는 에러인경우 빨강색으로
                tblResult.Text = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                If tblResult.Text = IISERROR Then
                    tblResult.ForeColor = vbRed
                Else
                    tblResult.ForeColor = vbBlack
                End If
        Next i
    End With
End Sub

Private Sub tblResult_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim strInfo As String       '수신한 추가정보
    
    If Row = 0 Then Exit Sub
    With tblResult
        .Row = Row: .Col = TResultEnum.ccInfo
        strInfo = .Text
        If Trim(strInfo) = "" Then Exit Sub
        
        MultiLine = 1
        TipWidth = 6000
        TipText = strInfo
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
End Sub

Private Sub MSComm_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long
    
    Select Case MSComm.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long
            
            Buffer = MSComm.Input
'            Buffer = "SMP_EDIT_DATA&aESAMPLErSEQ1800rDATE07Feb12rTIME13:07rDEVICESYRINGEmpH7.442mPCO213.3mmHgLmHmHct21%UNCORLcHCO3act8.8mmol/LcHCO3std13.5mmol/LctCO29.2mmol/LcO2SAT98.0%cBE(vt)-13.7mmol/LctHb(est)7.1g/dLmBP762mmHg&85"
'            Buffer = "SMP_EDIT_DATA&aESAMPLErSEQ1800rDATE07Feb12rTIME13:07rDEVICESYRINGEmpH7.442mPCO213.3mmHgLmPO213.3mmHgLmHmHct21%UNCORLcHCO3act8.8mmol/LcHCO3std13.5mmol/LctCO29.2mmol/LcO2SAT98.0%cBE(vt)-13.7mmol/LctHb(est)7.1g/dLmBP762mmHg&85"
            
'            Buffer = Buffer & "         " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "epoc BGEM Blood Test" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Patient ID: 12000102664" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Date & Time: 15-Nov-12 05:21:24" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Results: Gases+" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "pH       7.408                " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "pCO2     34.4   mmHg   Low    " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "pO2      112.9  mmHg   High   " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "cHCO3-   21.7   mmol/L        " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "BE(ecf)  -3.0   mmol/L Low    " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "cSO2     98.5   %      High   " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Results: Chem+" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Na+      139    mmol/L        " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "K+       4.5    mmol/L        " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Ca++     1.14   mmol/L Low    " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "cTCO2    22.7   mmol/L        " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Hct      36     %      Low    " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "cHgb     12.1   g/dL          " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "BE(b)    -2.4   mmol/L Low    " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Results: Meta+" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Glu      178    mg/dL  High   " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Lac      1.69   mmol/L High   " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Reference Ranges" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "pCO2       35.0 - 48.0   mmHg" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "pO2        83.0 - 108.0  mmHg" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "BE(ecf)    -2.0 - 3.0    mmol/L" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "cSO2       94.0 - 98.0   %" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Ca++       1.15 - 1.33   mmol/L" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Hct          38 - 51     %" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "BE(b)      -2.0 - 3.0    mmol/L" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Glu          74 - 100    mg/dL" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Lac        0.56 - 1.39   mmol/L" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Sample type: Unspecified" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Hemodilution: No" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Comments: " & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Operator: 1115" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Card lot: 07-12254-00" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Last EQC: 15-Nov-12 05:08:05" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Reader: 05928 (2.2.8.1)" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Host: 11035F5 (3.13.4)" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "Sensor config: 18.4" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "----------------------" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
'            Buffer = Buffer & "" & vbCrLf
            
            
            
            Call mIntLib.WriteLog(Buffer, ccEqp)
            
            lngBufLen = Len(Buffer)
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                
                Select Case mIntLib.Phase
                    Case 1      '## STX 대기
                        Select Case BufChar
                            Case STX
                                Call mIntLib.ClearBuffer
                                mIntLib.Phase = 2
                        End Select
                    Case 2      '## ETX 대기
                        Select Case BufChar
                            Case Chr(29) 'ETX
                                mIntLib.Phase = 3
                            Case Else
                                Call mIntLib.AddBuffer(BufChar)
                        End Select
                    Case 3      '## EOT 대기
                        Call EditRcvData
                        mIntLib.Phase = 1
                End Select
            Next i
        Case comEvSend
        
        Case comEvCTS
            EVMsg$ = "CTS 변경 감지"
        Case comEvDSR
            EVMsg$ = "DSR 변경 감지"
        Case comEvCD
            EVMsg$ = "CD 변경 감지"
        Case comEvRing
            EVMsg$ = "전화 벨이 울리는 중"
        Case comEvEOF
            EVMsg$ = "EOF 감지"

        '오류 메시지
        Case comBreak
            ERMsg$ = "중단 신호 수신"
        Case comCDTO
            ERMsg$ = "반송파 검출 시간 초과"
        Case comCTSTO
            ERMsg$ = "CTS 시간 초과"
        Case comDCB
            ERMsg$ = "DCB 검색 오류"
        Case comDSRTO
            ERMsg$ = "DSR 시간 초과"
        Case comFrame
            ERMsg$ = "프레이밍 오류"
        Case comOverrun
            ERMsg$ = "패리티 오류"
        Case comRxOver
            ERMsg$ = "수신 버퍼 초과"
        Case comRxParity
            ERMsg$ = "패리티 오류"
        Case comTxFull
            ERMsg$ = "전송 버퍼에 여유가 없음"
        Case Else
            ERMsg$ = "알 수 없는 오류 또는 이벤트"
    End Select

    If Len(EVMsg$) Then
        StatusBar.Panels(2).Text = EVMsg$
    ElseIf Len(ERMsg$) Then
        StatusBar.Panels(2).Text = ERMsg$
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부터 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim objIntInfo   As clsIISIntInfo    '인터페이스 검체정보 클래스
    Dim objIntNm     As clsIISIntNm      '장비별 검사항목 클래스
    
    Dim vWorkNo      As Variant  'Spread의 WorkNo
    Dim vBarNo       As Variant  'Spread의 바코드번호
    Dim strRcvBuf    As String   '수신한 Data
    Dim strIDRecord  As String   '수신한 Identifyer Record
    Dim strBarNo     As String   '수신한 BarNO
    Dim strWorkNo    As String   '수신한 WorkNo
    Dim strIntBase   As String   '장비기준 검사명
    Dim strIntResult As String   '수신한 검사결과
    Dim strResult    As String   'LIS결과
    Dim strFlags     As String   '수신한 Exception Group
    Dim strInfo      As String   '수신한 추가정보
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim Pos1         As Long
    Dim Pos2         As Long
    Dim i            As Long
    Dim ii           As Long
    Dim iCnt         As Long
    Dim varTmp       As Variant
    Dim intRow       As Long
    Dim blnSameBar   As Boolean
    
    strRcvBuf = mIntLib.Buffers(1).Buffers
    strRcvBuf = Replace(strRcvBuf, vbLf, "")
    varTmp = Split(strRcvBuf, vbCr)
    
    
    'H|\^&|164852||cobas ENERGIUM ^Roche Diagnostics|D4||||||P||20200827105539
    'M|1|EQU^RO^cobas ENERGIUM ^2.0|A32|20200827105539|5177
    'M|2|SAC^RO^cobas ENERGIUM ^2.0|||O371U0RB0|O371U0RB0|||20200827105516|I|||||5177|2
    'L|1|N
    
    For i = 0 To UBound(varTmp)
        strTemp1 = mGetP(varTmp(i), 1, "|")
        Select Case strTemp1
        Case "H"
        Case "P"
        Case "M"
            If mGetP(varTmp(i), 2, "|") = "2" Then
                strBarNo = mGetP(varTmp(i), 2, "|")
            End If
        Case "R"
            strResult = "SEEN"
        End Select
    Next
    
    If strBarNo <> "" And IsNumeric(strBarNo <> "") And Len(strBarNo) = "11" Then
        Call GetOrder(Trim(strBarNo))
        tblComplete.Row = tblComplete.DataRowCnt
        tblComplete.Col = 14
        tblComplete.Text = strResult
    End If

End Sub

'-----------------------------------------------------------------------------'
'   기능 : 결과판정, 결과저장, 화면표시
'   인수 :
'       - pIntInfo : 인터페이스 검체정보 클래스
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pIntInfo As clsIISIntInfo)
    Dim objAccInfo  As clsIISAccInfo     '접수내역 클래스
    Dim vBarNo      As Variant 'Spread의 바코드번호
    Dim strBarNo    As String  '바코드번호
    Dim strSpcYy    As String  '검체연도
    Dim lngSpcNo    As Long    '검체번호
    Dim i           As Long
    
    Me.MousePointer = vbHourglass
    
    strBarNo = pIntInfo.BarNo
    
    '## 결과판정
    If mIntLib.CheckResult(pIntInfo) = -1 Then
        '## 접수정보가 없을때 결과표시
        Call SetComplete1(pIntInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        '## 접수정보가 있을때 결과표시
        strSpcYy = Mid$(strBarNo, 1, SPCYYLEN)
        lngSpcNo = CLng(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
        Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
        
        Call SetComplete2(objAccInfo, pIntInfo.SpcPos)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Set objAccInfo = Nothing
        
        '## ClientDb, Server에 결과저장
        Call mIntLib.SaveClientDb(strSpcYy, lngSpcNo)
        Call mIntLib.SaveResult(strSpcYy, lngSpcNo)
        Call mIntLib.Remove(strSpcYy, lngSpcNo)
        StatusBar.Panels(2).Text = "검체번호:" & strBarNo & " 를 정상적으로 결과저장 했습니다."
    End If
    
    '## tblReady에서 전송된 검체삭제
    If mIntLib.BarPos = ccPC Then
        With tblReady
            For i = 1 To .DataRowCnt
                Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
                If CStr(vBarNo) = strBarNo Then
                    Call .DeleteRows(i, 1)
                    Exit For
                End If
            Next i
        End With
    End If

    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    
    If pBarNo = "" Then Exit Sub
'    pBarNo = Mid(pBarNo, 2)
'    Set objAccInfo = mIntLib.GetAccInfo(pBarNo, chkPOC.Value)
    Set objAccInfo = mIntLib.GetAccInfo_New(pBarNo, chkPOC.Value)
    If Not (objAccInfo Is Nothing) Then
        '## tblReady, tblresult, Label에 정보표시
        Call SetReady(objAccInfo)
        Call SetLabel(objAccInfo)
        Call SetResult(objAccInfo)
        
        Set objAccInfo = Nothing
    End If
    'txtBarNo.Text = "": txtBarNo.SetFocus
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblReady에 정보표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetReady(ByVal pAccInfo As clsIISAccInfo)
    Dim lngWorkNo As Long   'WorkNo
    
    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If
        
        '## WorkNo 구하기
        If .DataRowCnt = 0 Then
            If Trim(txtWorkNo.Text) <> "" Then
                lngWorkNo = CLng(txtWorkNo.Text)
                txtWorkNo.Text = CStr(lngWorkNo + 1)
            Else
                lngWorkNo = 1
                txtWorkNo.Text = CStr(lngWorkNo + 1)
            End If
        Else
            lngWorkNo = CLng(txtWorkNo.Text)
            txtWorkNo.Text = CStr(lngWorkNo + 1)
        End If
        
        .Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
        .Col = TReadyEnum.ccBarNo:  .Text = pAccInfo.GetBarNo
        .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.WorkArea, pAccInfo.AccDt, pAccInfo.AccSeq)
        
        If pAccInfo.QcFg = "0" Then         '## 일반검체
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
        ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
        End If
        Call .SetActiveCell(1, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblComplete에 정보표시 (접수정보가 없을때)
'   인수 :
'       - pIntInfo : 인터페이스 검체정보 클래스
'-----------------------------------------------------------------------------'
Private Sub SetComplete1(ByVal pIntInfo As clsIISIntInfo)
    Dim objIntResult As clsIISIntResult     '인터페이스 결과 클래스
    Dim i            As Long
    
    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TCompleteEnum.ccNo:      .Text = pIntInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pIntInfo.BarNo
        .Col = TCompleteEnum.ccSendCnt: .Text = pIntInfo.IntResults.Count
        
        For Each objIntResult In pIntInfo.IntResults
            If .MaxCols <= .DataColCnt Then
                .MaxCols = .MaxCols + 1
            End If
            .Col = TCompleteEnum.ccResult + i
            .ColHidden = True
            .Text = objIntResult.IntNm & DIV & objIntResult.Result & DIV & DIV & DIV & DIV & _
                    DIV & DIV & objIntResult.Info
            i = i + 1
        Next
        Set objIntResult = Nothing
        
        Call .SetActiveCell(TCompleteEnum.ccNo, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblComplete에 정보표시 (접수정보가 있을때)
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetComplete2(ByVal pAccInfo As clsIISAccInfo, ByVal strSpcPos As String)
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim i           As Long
    
    With tblComplete
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TCompleteEnum.ccNo:      .Text = strSpcPos 'pAccInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pAccInfo.GetBarNo
        .Col = TCompleteEnum.ccAccNo:   .Text = mGetAccNo(pAccInfo.WorkArea, pAccInfo.AccDt, pAccInfo.AccSeq)
        
        If pAccInfo.QcFg = "0" Then         '## 일반검체
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.PtId
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.Name
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.Sex & " / " & pAccInfo.Age
            .Col = TCompleteEnum.ccDoctNm:  .Text = pAccInfo.OrdDoctNm
            .Col = TCompleteEnum.ccDeptNm:  .Text = pAccInfo.DeptNm
            .Col = TCompleteEnum.ccWardNm:  .Text = pAccInfo.WardNm
            .Col = TCompleteEnum.ccStatFg:  .Text = IIf(pAccInfo.StatFg = "1", "Y", "N")
            .Col = TCompleteEnum.ccSpcNm:   .Text = pAccInfo.SpcNm
            .Col = TCompleteEnum.ccQcFg:    .Text = pAccInfo.QcFg
            .Col = TCompleteEnum.ccSendCnt: .Text = pAccInfo.SendCnt
            
            For Each objResult In pAccInfo.Results
                If .MaxCols <= .DataColCnt Then
                    .MaxCols = .MaxCols + 1
                End If
                .Col = TCompleteEnum.ccResult + i
                .ColHidden = True
                .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                        DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & DIV & _
                        IIf(objResult.Ref.RefFg = "1", mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal), "") & DIV & _
                        objResult.IntInfo
                i = i + 1
            Next
            Set objResult = Nothing
        ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.CtrlCd
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.LevelCd
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.LotNo
            .Col = TCompleteEnum.ccQcFg:    .Text = pAccInfo.QcFg
            .Col = TCompleteEnum.ccSendCnt: .Text = pAccInfo.SendCnt
            
            For Each objQCResult In pAccInfo.QCResults
                If .MaxCols <= .DataColCnt Then
                    .MaxCols = .MaxCols + 1
                End If
                .Col = TCompleteEnum.ccResult + i
                .ColHidden = True
                .Text = objQCResult.IntNm.IntNm & DIV & objQCResult.IntResult & DIV & _
                        objQCResult.RstCd & DIV & objQCResult.Unit & DIV & objQCResult.RADiv
                i = i + 1
            Next
            Set objQCResult = Nothing
        End If
        Call .SetActiveCell(TCompleteEnum.ccNo, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblResult 정보표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetResult(ByVal pAccInfo As clsIISAccInfo)
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    
    Call mTblClear(tblResult)
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        For Each objResult In pAccInfo.Results
            With tblResult
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = TResultEnum.ccTestNm:    .Text = objResult.IntNm.IntNm
                .Col = TResultEnum.ccEqpResult: .Text = objResult.Result
                .Col = TResultEnum.ccLISResult: .Text = objResult.RstCd
                .Col = TResultEnum.ccUnit:      .Text = objResult.Unit
                .Col = TResultEnum.ccHLDiv:     .Text = objResult.HLDiv
                .Col = TResultEnum.ccDPDiv:     .Text = objResult.DPDiv
                .Col = TResultEnum.ccRef
                .Text = IIf(objResult.Ref.RefFg = "1", mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal), "")
            End With
        Next
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        For Each objQCResult In pAccInfo.QCResults
            With tblResult
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                .Col = TResultEnum.ccTestNm:    .Text = objQCResult.IntNm.IntNm
                .Col = TResultEnum.ccEqpResult: .Text = objQCResult.Result
                .Col = TResultEnum.ccLISResult: .Text = objQCResult.RstCd
                .Col = TResultEnum.ccHLDiv:     .Text = objQCResult.RADiv
            End With
        Next
    End If
    
    Set objResult = Nothing
    Set objQCResult = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Label에 환자정보, 접수정보 표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetLabel(ByVal pAccInfo As clsIISAccInfo)
    Call CtlClear(ccLabel)
    
    If pAccInfo.QcFg = "0" Then         '## 일반검체
        Call LabelShow("0")
        lblPtId.Caption = pAccInfo.PtId
        lblName.Caption = pAccInfo.Name
        lblSexAge.Caption = pAccInfo.Sex & " / " & pAccInfo.Age
        lblDoctNm.Caption = pAccInfo.OrdDoctNm
        lblDeptNm.Caption = pAccInfo.DeptNm
        lblWardNm.Caption = pAccInfo.WardNm
        lblStatFg.Caption = IIf(pAccInfo.StatFg = "1", "Y", "N")
        lblSpcNm.Caption = pAccInfo.SpcNm
    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
        Call LabelShow("1")
        lblPtId.Caption = pAccInfo.CtrlCd
        lblName.Caption = pAccInfo.LevelCd
        lblSexAge.Caption = pAccInfo.LotNo
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 검체종류에 따라 Label을 다르게 표시
'   인수 :
'       - pQcFg : 0(일반검체), 1(QC검체)
'-----------------------------------------------------------------------------'
Private Sub LabelShow(ByVal pQcFg As String)
    Dim i As Long
    
    If pQcFg = "0" Then         '## 일반검체
        lblControl.Caption = "환  자 ID :"
        lblLevel.Caption = "이     름 :"
        lblLotNo.Caption = "성별/나이 :"
        For i = 0 To lblGeneral.Count - 1
            lblGeneral(i).Visible = True
        Next i
        
        lblDoctNm.Visible = True:   lblDeptNm.Visible = True
        lblWardNm.Visible = True:   lblStatFg.Visible = True
        lblSpcNm.Visible = True
    ElseIf pQcFg = "1" Then     '## QC검체
        lblControl.Caption = "Control :"
        lblLevel.Caption = "Level   :"
        lblLotNo.Caption = "Lot No  :"
        For i = 0 To lblGeneral.Count - 1
            lblGeneral(i).Visible = False
        Next i
        
        lblDoctNm.Visible = False:   lblDeptNm.Visible = False
        lblWardNm.Visible = False:   lblStatFg.Visible = False
        lblSpcNm.Visible = False
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비설정 정보조회, 포트 Open
'-----------------------------------------------------------------------------'
Private Sub GetEqpComm()
    Dim objComm     As clsIISEqpComm    '통신설정 클래스
    Dim strErrMsg   As String           '에러메시지

    '## 통신설정 정보조회
    Set objComm = mIntLib.GetEqpComm
    If objComm Is Nothing Then Exit Sub

    With objComm
        MSComm.CommPort = .Port
        MSComm.Settings = .GetSettings
    End With
    Set objComm = Nothing

On Error GoTo Errors
    '## 포트 Open
    With MSComm
        '## 이미 포트가 열린경우
        If .PortOpen Then
            strErrMsg = mEqpCd & " 장비의 통신포트가 이미 열려있습니다."
            Error.SetLog App.EXEName, "frmIISRapidLab348_1", "GetEqpComm", strErrMsg, Now
            Call mIntLib_EqpError("E004")
            Exit Sub
        End If

        .RThreshold = 1
        .SThreshold = 1
        .RTSEnable = True
        .PortOpen = True

    End With

    '## 보관일이 지난데이터 삭제
    Call mIntLib.DelHistoryData
    Exit Sub

Errors:
    '## 다른 장치에서 포트를 사용하는 경우
    If Err.Number = 8005 Then
        strErrMsg = mEqpCd & " 장비에 설정된 포트가 이미 사용중입니다."
        Error.SetLog App.EXEName, "frmIISRapidLab348_1", "GetEqpComm", strErrMsg, Now
        Call mIntLib_EqpError("E005")
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 수신한 Exception에 대한 정보 조회
'   인수 :
'       - pFlags  : 수신한 Exception Group
'       - pResult : 수신한 결과
'   반환 : Exception에 대한 상세정보
'-----------------------------------------------------------------------------'
Private Function GetIntInfo(ByVal pFlags As String, ByRef pResult As String) As String
    Dim aryFlags()  As String   'Mnemonic 배열
    Dim strMeaning  As String   'Meaning
    Dim strTemp     As String
    Dim i           As Long
    
    If Trim$(pFlags) = "" Then Exit Function
    
    aryFlags = Split(pFlags, ETB)
    For i = LBound(aryFlags) To UBound(aryFlags)
        If Trim$(aryFlags(i)) = "" Then Exit For
        
        If i = 0 Then
            strTemp = Space(2) & "[Mnemonic]: " & aryFlags(i) & vbCrLf
        Else
            strTemp = strTemp & vbCrLf & Space(2) & "[Mnemonic]: " & aryFlags(i) & vbCrLf
        End If
        
        strMeaning = ""
        Select Case aryFlags(i)
            Case "COOXERR": strMeaning = "error that prevents measurement of CO-ox module parameters"
            Case "DRIFT":   strMeaning = "driift (D2)"
            Case "<":       strMeaning = "below reporting range"
            Case ">":       strMeaning = "above reporting range"
            Case "QV":      strMeaning = "if blood, questionable CO-oximeter data"
            Case "HBF":     strMeaning = "fetal hemoglobin (corrected) sample"
            Case "H":       strMeaning = "above reference or expected range"
            Case "HH":      strMeaning = "above action range"
            Case "L":       strMeaning = "below reference or expected range"
            Case "LL":      strMeaning = "below action range"
            Case "NEP":     strMeaning = "no end point"
            Case "INTERF":  strMeaning = "glucose/lactate interferant detected"
            Case "QUES":    strMeaning = "noise"
            Case "BUB"
                '## BUG일때만 Error 표시
                strMeaning = "bubbles or short sample detected"
                pResult = IISERROR
            Case "TEMP":    strMeaning = "temperature warning"
            Case "CTEMP":   strMeaning = "CO-ex module sample chamber out of temperature range"
            Case "CINTERF": strMeaning = "CO-ex module interferent detected"
            Case "SULF":    strMeaning = ">1.5% SulfHb detected"
        End Select
        
        If strMeaning <> "" Then
            strTemp = strTemp & Space(2) & "[Meaning ]: " & strMeaning & vbCrLf
        End If
    Next i
    
    GetIntInfo = vbCrLf & strTemp
End Function

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear(Optional ByVal pFlag As ClearEnum = ccAll)
    lblPtId.Caption = "":       lblName.Caption = ""
    lblSexAge.Caption = "":     lblDoctNm.Caption = ""
    lblDeptNm.Caption = "":     lblWardNm.Caption = ""
    lblStatFg.Caption = "":     lblSpcNm.Caption = ""
    
    If pFlag = ccAll Then
        txtBarNo.Text = "":         Call mTblClear(tblResult)
        Call mTblClear(tblReady):   Call mTblClear(tblComplete)
    End If
End Sub

'------------------------------------------------------------------'
'   기능 : 장비설정 관련 에러처리
'------------------------------------------------------------------'
Private Sub mIntLib_EqpError(ByVal pCode As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey)
End Sub

'------------------------------------------------------------------'
'   기능 : 검체관련 에러처리1
'------------------------------------------------------------------'
Private Sub mIntLib_SpcError(ByVal pCode As String, ByVal pBarNo As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo)
End Sub

'------------------------------------------------------------------'
'   기능 : 검체관련 에러처리2
'------------------------------------------------------------------'
Private Sub mIntLib_SpcErrorX(ByVal pCode As String, ByVal pBarNo As String, ByVal pPtId As String, ByVal pName As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo, pPtId, pName)
End Sub

'------------------------------------------------------------------'
'   기능 : Popup 메뉴 Click 이벤트
'------------------------------------------------------------------'
Private Sub mPopup_Click(ByVal vMenuID As Long)
    Dim vBarNo      As Variant  'Spread의 바코드번호
    Dim strSpcYy    As String   '검체연도
    Dim lngSpcNo    As Long     '검체번호
    
    Select Case vMenuID
        Case DELETE     '## Delete
            With tblReady
                Call .GetText(TReadyEnum.ccBarNo, .ActiveRow, vBarNo)
                If vBarNo <> "" Then
                    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
                    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
                    Call mIntLib.AccInfos.Remove(strSpcYy, lngSpcNo)
                    Call .DeleteRows(.ActiveRow, 1)
                    Call mTblClear(tblResult)
                End If
            End With
        Case DELETEALL  '## Delete All
            Call mIntLib.AccInfos.RemoveAll
            Call mTblClear(tblReady)
    End Select
End Sub




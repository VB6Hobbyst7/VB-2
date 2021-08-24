VERSION 5.00
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIISLiaison 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Liaison"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "굴림"
      Size            =   9.75
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3855
      Left            =   15990
      TabIndex        =   27
      Top             =   3420
      Visible         =   0   'False
      Width           =   7455
      Begin VB.TextBox txtOrder 
         Height          =   405
         Left            =   2970
         TabIndex        =   37
         Text            =   "MO31"
         Top             =   1740
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtBarNo 
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
         Left            =   2940
         TabIndex        =   35
         Text            =   "123456789011"
         Top             =   1260
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "조회"
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
         Left            =   4560
         TabIndex        =   32
         Top             =   510
         Width           =   1005
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "오더전송"
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
         Left            =   5610
         TabIndex        =   30
         Top             =   510
         Width           =   1005
      End
      Begin VB.TextBox txtTmrSec 
         Alignment       =   2  '가운데 맞춤
         Height          =   285
         Left            =   5820
         TabIndex        =   29
         Text            =   "30"
         Top             =   120
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CheckBox chkTmr 
         Caption         =   "자동오더"
         Height          =   225
         Left            =   4680
         TabIndex        =   28
         Top             =   180
         Value           =   1  '확인
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFrDate 
         Height          =   315
         Left            =   1320
         TabIndex        =   31
         Top             =   510
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   122224641
         CurrentDate     =   40270
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   2940
         TabIndex        =   33
         Top             =   510
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   122224641
         CurrentDate     =   40270
      End
      Begin FPSpread.vaSpread tblexcel 
         Height          =   675
         Left            =   2880
         TabIndex        =   38
         Top             =   2340
         Visible         =   0   'False
         Width           =   675
         _Version        =   393216
         _ExtentX        =   1191
         _ExtentY        =   1191
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
         SpreadDesigner  =   "frmIISLiaison.frx":0000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "바코드번호 : "
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
         Left            =   1650
         TabIndex        =   36
         Top             =   1335
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBE6E6&
         Caption         =   "▶ 접수일자"
         Height          =   180
         Left            =   240
         TabIndex        =   34
         Top             =   585
         Width           =   960
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   2670
      TabIndex        =   26
      Top             =   -60
      Visible         =   0   'False
      Width           =   675
   End
   Begin FPSpread.vaSpread tblReady 
      Height          =   3600
      Left            =   105
      TabIndex        =   21
      Top             =   525
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   6350
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
      MaxCols         =   6
      MaxRows         =   10
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISLiaison.frx":0203
   End
   Begin VB.Timer tmrResult 
      Left            =   5520
      Top             =   8520
   End
   Begin MSComDlg.CommonDialog AlloFile 
      Left            =   6000
      Top             =   8490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   6548
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   107
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
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12698
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11483
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13913
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8567
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1290
      Left            =   6548
      TabIndex        =   4
      Top             =   407
      Width           =   8595
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1245
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   20
         Top             =   600
         Width           =   900
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
         TabIndex        =   19
         Top             =   240
         Width           =   900
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
         TabIndex        =   18
         Top             =   975
         Width           =   810
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
         TabIndex        =   17
         Top             =   600
         Width           =   720
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
         TabIndex        =   16
         Top             =   240
         Width           =   720
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
         TabIndex        =   15
         Top             =   975
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
         TabIndex        =   14
         Top             =   600
         Width           =   990
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
         TabIndex        =   13
         Top             =   240
         Width           =   990
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6600
      Top             =   8430
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4110
      Left            =   105
      TabIndex        =   22
      Top             =   4725
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   7250
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
      MaxCols         =   14
      MaxRows         =   14
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISLiaison.frx":0738
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   105
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   105
      Width           =   2805
      _ExtentX        =   4948
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
      Caption         =   "■ 검사대상 리스트"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   405
      Left            =   105
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4305
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   714
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
      Caption         =   "■ 검사완료 리스트"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblResult 
      Height          =   6690
      Left            =   6555
      TabIndex        =   25
      Top             =   1710
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
      SpreadDesigner  =   "frmIISLiaison.frx":0F98
   End
   Begin MDIControls.MDIActiveX MDIActiveX 
      Left            =   7230
      Top             =   8550
      _ExtentX        =   847
      _ExtentY        =   794
   End
End
Attribute VB_Name = "frmIISLiaison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISLiaison.frm
'   작성자  : 오세원
'   내  용  : Liaison 장비폼
'   작성일  : 2018-01-08
'   버  전  : 1.0.0
'   병  원  :
'       1. 전주예수병원
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady의 Column Enum
Private Enum TReadyEnum
    ccNo = 1
    ccBarNo = 2
    ccAccNo = 3
    ccPtId = 4
    ccName = 5
    ccOrder = 6
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
    ccClass = 8
    '-- 2015.08.28 추가
    ccIntBase = 9
End Enum

'## Clear Enum
Private Enum ClearEnum
    ccAll = 1
    ccLabel = 2
End Enum

'## Popup Menu ID
Private Const DELETE    As Long = 1
Private Const DELETEALL As Long = 2

Private WithEvents mIntLib  As clsIISInterface   '인터페이스 클래스
Attribute mIntLib.VB_VarHelpID = -1
Private WithEvents mPopup   As clsIISPopup       '팝업메뉴
Attribute mPopup.VB_VarHelpID = -1

Private mIntErrors  As clsIISIntErrors           '인터페이스 에러 컬렉션
Private mOrder      As clsIISIntOrder           '오더정보 클래스

Private mEqpCd  As String   '장비코드
Private mEqpKey As String   '장비키

'임시 - 데모용
Dim strTransData    As Variant

Dim AdoCn           As ADODB.Connection
Dim AdoRS           As ADODB.Recordset

Private Const AM As Variant = -0.696669715055768
Private Const AN As Variant = 3.57268581287044
Private Const BM As Variant = -0.58804147563697
Private Const BN As Variant = 3.68672614195763
Private Const CM As Variant = -0.424608277908711
Private Const Cn As Variant = 3.85830192881032

Private Const VarA As Variant = 0
Private Const VarD As Variant = 150

Private gBarNo  As String


Dim strBuffer As String
Dim intPatSeq As Integer

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property


Public Function ReadFileBinary(ByVal strFileName As String) As String

On Error GoTo errHandler
    Dim fsT, tFilePath As String

    Set fsT = CreateObject("ADODB.Stream")

    fsT.Type = 2
 
    fsT.Charset = "utf-8"
 
    fsT.Open
'    fsT.Type = adTypeBinary
'    fsT.Type = adTypeText
    fsT.LoadFromFile strFileName
   
    Dim strText As String
    strText = ""
    Do Until fsT.EOS
        strText = strText & fsT.ReadText(adReadLine) & vbLf     ' 줄바꿈 추가
    Loop
 

    fsT.Close

    ReadFileBinary = strText
    GoTo finish
 
errHandler:

    MsgBox (Err.Description)
    Exit Function

finish:

End Function

Public Sub FileUp(PropNo As Long, FilePath As String, FileName As String)
    Dim Cn As ADODB.Connection
    Dim Cmd As ADODB.Command
    Dim Object As ADODB.Parameter
'    Dim Fso As FileSystemObject
    Dim Chunk() As Byte
    Dim Sql As String
    Dim FileType As String  '파일확장자를 저장.
    Dim Fd As Integer       '파일핸들
    Dim Flen As Long
    Dim szConn As String
    
    On Error GoTo ErrLog
    
    Dim Rs As ADODB.Recordset
    
    'Set Fso = Nothing
    Set Object = Nothing
    Set Cmd = Nothing
    Set Cn = Nothing
    
    
    'Up은 한글지원문제로 인하여 ODBC에서 Oracle ODBC로 설정한 후
    'OleDB For ODBC를 사용하여 연결한다.
    szConn = "Provider=MSDAORA.1;Data Source=Liaison"

'        "Dbq=" & Mid(Dialog1.FileName, 1, InStrRev(Dialog1.FileName, "\")) & ";" & _

szConn = "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
        "Dbq=C:\;" & _
        "Extensions=asc,csv,tab,txt;       "




    Set Cn = New Connection '커넥션개체 생성
    Set Cmd = New Command   '커맨드캐체 생성
'
    Cn.ConnectionString = szConn
    
    Cn.Open
    
    Set Rs = New ADODB.Recordset
    
    '해당하는 PropNo의 파일을 Update
    Sql = "SELECT * FROM export.csv;"
    Rs.Open Sql, Cn ', adOpenDynamic, adLockReadOnly ', adCmdText

            Dim i As Integer
            
    Do Until Rs.EOF
        For i = 1 To 20
        Debug.Print Rs.Fields(i).Value
        Next
        Rs.MoveNext
        
    Loop
    
    Rs.Close
    
    '파일확장자는 마침표를 포함하여 최대 5자리를 허용한다.
    FileType = Mid(FileName, InStrRev(FileName, "."))
    FileType = Trim(Left(FileType, 5))
    
    With Cmd
        .ActiveConnection = Cn
        .CommandText = Sql
        .CommandType = adCmdText

        .Parameters.Append .CreateParameter("@file_size", adInteger, adParamInput, , FileLen(FilePath & FileName))
        .Parameters.Append .CreateParameter("@file_type", adVarChar, adParamInput, 5, FileType)
        
    End With
    
    '파일을 이전파일로 오픈하여 읽는다.
    Fd = FreeFile

    Open FilePath & FileName For Binary Access Read As Fd

    Flen = LOF(Fd)

    '파일크기가 0일때
    If Flen = 0 Then
        Close Fd
        Set Cmd = Nothing
        Set Cn = Nothing

        MsgBox "Error while opening the file"
        Exit Sub
    End If

    Set Object = Cmd.CreateParameter("object", adLongVarBinary, adParamInput, Flen + 100)

    ReDim Chunk(1 To Flen)
    Get Fd, , Chunk()

    Object.AppendChunk Chunk()
    Cmd.Parameters.Append Object

    Close Fd
        
    Cmd.Execute
        
    Cn.Close
    
'    Set Fso = New FileSystemObject
    
'    '폴더에 파일을 BLOB 테이블로 올린후 삭제한다.
'    If Fso.FileExists(FilePath & FileName) Then
'        Fso.DeleteFile (FilePath & FileName)
'    End If
    
'    Set Fso = Nothing
    Set Object = Nothing
    Set Cmd = Nothing
    Set Cn = Nothing
    
    Exit Sub
    
ErrLog:
    'WriteEventLog "FileUp(" & PropNo & "," & FilePath & "," & FileName & ") " & vbCrLf & Err.Number & "  " & Err.Description
'    Set Fso = Nothing
    Set Object = Nothing
    Set Cmd = Nothing
    Set Cn = Nothing
End Sub

 

'
'Private Sub cmdGetRslt_Click()
''    Call Get_SearchList
'
'    Dim intRow      As Integer
'    Dim intCol      As Integer
'    Dim strTmp      As String
'    Dim intIdx      As Integer
'    Dim strSrcfile  As String
'    Dim strDestFile As String
'    Dim strBuffer   As String
'    Dim strtmpBuf   As String
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'    Dim intCnt      As Integer
'    Dim varTmp      As Variant
'
'    Dim fName As String
'    Dim Buf() As Byte
'    Dim r As Long
'
'
'    'Call FileUp(0, "", "")
'
'   'Call OpenExcel
'
'    'Call Excel_Open
'   ' strtmpBuf = ""
'
'
'   ' Dim strBuffer As String
'    Dim strBuf      As String
'    Dim strFile     As String
'    Dim varBuf As Variant
'
'    strFile = GetLiaisonConfig("ExportPath") & "\export.csv"
'
'    Open strFile For Input As #3
'
'    strBuffer = ""
'
'    strBuf = ReadFileBinary(strFile)
'    Close #3
'
'    With vasExcel
'        Call mIntLib.ClearBuffer
'        varBuf = Split(strBuf, vbLf)
'
'        For intRow = 0 To UBound(varBuf)
'            mIntLib.BufCnt = mIntLib.BufCnt + 1
'            Call mIntLib.AddBuffer(varBuf(intRow))
'            strtmpBuf = ""
'        Next
'    End With
'
'    Call EditRcvData
'
'    Call mIntLib.ClearBuffer
'
'Exit Sub
'
'ErrRoutine:
'
'
'End Sub

Private Sub cmdMakeWS_Click()
''    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
''    Dim objResult   As clsIISResult     '결과내역 클래스
''    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
''    Dim mLogOn As clsIISLogOn
'
''    Dim strAlloFile As String
'    Dim lngFIleNum  As Long
'    Dim strInFo     As String
''    Dim strOldInFo  As String
'
'
'    Dim iCnt As Integer
'    Dim varTmp As Variant
'    Dim strBarNo As String
'    Dim strLabNo As String
'    Dim strFNm   As String
'    Dim strLNm   As String
'    Dim strBirth As String
'    Dim strSex   As String
'    Dim strAge   As String
'
'    Dim intCnt As Integer
'
'    Screen.MousePointer = 11
'
'
'    With AlloFile
'        .CancelError = True
'        .FileName = GetLiaisonConfig("ImportPath") & "\" & Trim(txtFileNm.Text) & ".csv"
'        If Len(Dir(.FileName)) Then Kill .FileName
'
'        lngFIleNum = FreeFile
'
'        Open .FileName For Append As #lngFIleNum
'
'
'        strInFo = "Col1;Col2;Col3;Col4;Col5;Col6;Col7" & vbCrLf
'
'        intCnt = 1
'        For iCnt = 1 To tblReady.DataRowCnt
'            tblReady.GetText 2, iCnt, varTmp
'            strBarNo = varTmp
'            tblReady.GetText 4, iCnt, varTmp
'            strLabNo = varTmp
'            tblReady.GetText 5, iCnt, varTmp
'            strFNm = Mid(varTmp, 1, 1)
'            strLNm = Mid(varTmp, 2)
'            strFNm = Han2Eng.HanToEng(strFNm)
'            strLNm = Han2Eng.HanToEng(strLNm)
'            tblReady.GetText 6, iCnt, varTmp
'            varTmp = Split(varTmp, "|")
'            strSex = varTmp(0)
'            If strSex = "M" Then
'                strSex = "Male"
'            Else
'                strSex = "FeMale"
'            End If
'
'            strAge = varTmp(1)
'            If Len(strAge) = 13 Then
'                If Mid(strAge, 7, 1) = "1" Or Mid(strAge, 7, 1) = "2" Then
'                    strAge = "19" & Mid(strAge, 1, 6)
'                ElseIf Mid(strAge, 7, 1) = "3" Or Mid(strAge, 7, 1) = "4" Then
'                    strAge = "20" & Mid(strAge, 1, 6)
'                Else
'                    strAge = ""
'                End If
'            Else
'                strAge = ""
'            End If
'
'            If strAge <> "" Then
'                strAge = Mid(strAge, 1, 4) & "." & Mid(strAge, 5, 2) & "." & Mid(strAge, 7, 2)
'            End If
'
'                                                                                        'BioSciTec / Panel / Rev. 001;Sample 3;345;Kant;Immanuel;Male;1804.02.12
'
'
'            tblReady.GetText 1, iCnt, varTmp
'            If varTmp = "IN" Then
'                strInFo = strInFo & "Liaison / Panel 30 KO Inhalant A / Rev. 008;" & strBarNo & ";" & strLabNo & ";" & strLNm & ";" & strFNm & ";" & strSex & ";" & strAge & vbCrLf
'                If iCnt = tblReady.DataRowCnt Then
'                    strInFo = strInFo & "Liaison / Panel 30 KO Inhalant B / Rev. 008;" & strBarNo & ";" & strLabNo & ";" & strLNm & ";" & strFNm & ";" & strSex & ";" & strAge & vbCrLf
'                Else
'                    strInFo = strInFo & "Liaison / Panel 30 KO Inhalant B / Rev. 008;" & strBarNo & ";" & strLabNo & ";" & strLNm & ";" & strFNm & ";" & strSex & ";" & strAge & vbCrLf
'                End If
'            ElseIf varTmp = "FD" Then
'                strInFo = strInFo & "Liaison / Panel 30 KO Food  A / Rev. 004;" & strBarNo & ";" & strLabNo & ";" & strLNm & ";" & strFNm & ";" & strSex & ";" & strAge & vbCrLf
'                If iCnt = tblReady.DataRowCnt Then
'                    strInFo = strInFo & "Liaison / Panel 30 KO Food  B / Rev. 009;" & strBarNo & ";" & strLabNo & ";" & strLNm & ";" & strFNm & ";" & strSex & ";" & strAge & vbCrLf
'                Else
'                    strInFo = strInFo & "Liaison / Panel 30 KO Food  B / Rev. 009;" & strBarNo & ";" & strLabNo & ";" & strLNm & ";" & strFNm & ";" & strSex & ";" & strAge & vbCrLf
'                End If
'            End If
'
'
'            intCnt = intCnt + 2
'
'
'        Next
'       ' Call vasExcel.ExportToExcel(.FileName, txtFileNm.Text, "c:\1.txt")
'
'        If strInFo <> "" Then
'            strInFo = Mid(strInFo, 1, Len(strInFo) - 2)
'        End If
'
'        Print #lngFIleNum, strInFo
'        Close #lngFIleNum
'    End With
'
'    MsgBox "워크리스트 생성이 완료되었습니다", vbInformation + vbOKOnly, Me.Caption
'    Screen.MousePointer = 0

End Sub

Private Function GetLiaisonConfig(ByVal strConfigNm As String) As String

Dim strFileName As String
Dim strReturnedString As String

    strFileName = App.Path & "\Liaison.ini"
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "Liaison", strConfigNm, "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    GetLiaisonConfig = strReturnedString
    
End Function


Private Sub chkTmr_Click()
    
    If chkTmr.Value = "1" Then
        tmrResult.Interval = txtTmrSec.Text * 1000 ' 30000
        tmrResult.Enabled = True
    Else
        tmrResult.Enabled = False
    End If
    
End Sub

Private Sub cmdSearch_Click()
    Dim varTemp As Variant
    Dim i       As Integer
    Dim strBarcode  As String
    Dim intRow  As Integer
    Dim blnSame As Boolean
    
    varTemp = Get_NewResult
    varTemp = Split(varTemp, "/")
    blnSame = False
    
    For i = 0 To UBound(varTemp)
        If varTemp(i) <> "" Then
            strBarcode = varTemp(i)
            With tblReady
                For intRow = 1 To .DataRowCnt
                    .Row = intRow
                    .Col = TReadyEnum.ccNo
                    If Trim(.Text) <> "SEND" Then
                        .Col = TReadyEnum.ccBarNo
                        If Mid(Trim(.Text), 1, 11) = Mid(strBarcode, 1, 11) Then
                            'strBarcode = strBarcode & CheckDisit(strBarcode)
                            'Call GetOrder(strBarcode)
                            blnSame = True
                        Else
                            blnSame = False
                        End If
                    End If
                Next
                
                If blnSame = False Then
                    strBarcode = strBarcode & CheckDisit(strBarcode)
                    Call GetOrder(strBarcode)
                End If
                
            End With
            
        End If
    Next
    
    
End Sub

'   Newest Result Recordset
Public Function Get_NewResult() As String

    Dim Rs          As ADODB.Recordset
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim strEqpCd    As String           '장비코드
    Dim strFromDt   As String           'From Date
    Dim strToDt     As String           'To Date
    Dim strKey      As String           'Spread의 키(SpcYy+SpcNo)
    Dim strSpcYy    As String           '바코드번호(연도)
    Dim strSpcNo    As String           '바코드번호(순번)
    Dim strTemp     As String
    Dim varTemp     As Variant
    Dim i           As Integer
    
    strFromDt = Format$(dtpFrDate.Value, "YYYYMMDD")
    strToDt = Format$(dtpToDate.Value, "YYYYMMDD")
    
    Me.MousePointer = vbHourglass
    'Call mTblClear(tblReady)
    strTemp = ""
    
On Error GoTo Errors
    Set objAccInfo = New clsIISAccInfo
    'Set Rs = objAccInfo.GetTargetSpcs(mEqpCd, strFromDt, strToDt)
    Set Rs = objAccInfo.GetTargetSpcs_ABO(mEqpCd, strFromDt, strToDt)
    If Not (Rs.BOF Or Rs.EOF) Then
        With tblReady
            Do Until Rs.EOF
                strSpcYy = Rs.Fields("SPCYY").Value
                strSpcNo = Rs.Fields("SPCNO").Value
        
                strKey = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(strSpcNo, String$(SPCNOLEN, "0"))
                
                'strKey = strSpcYy & strSpcNo
                If strTemp <> strKey Then
                    '## 다른 바코드번호 일때는 모든정보 표시
                    varTemp = varTemp & strKey & "/"
                    
                Else
                    '## 같은 바코드번호 일때는 검사명만 표시
'                    .Col = TReadyEnum.ccTestNms
'                    .Value = .Value & "," & Rs.Fields("TESTNM").Value & ""
                End If
                strTemp = strKey
                Rs.MoveNext
            Loop
            
        End With
    End If

    Rs.Close
    Set Rs = Nothing
    Set objAccInfo = Nothing
    Me.MousePointer = vbDefault
    
    Get_NewResult = varTemp
    
    Exit Function
    
Errors:
    Set Rs = Nothing
    Set objAccInfo = Nothing
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "오류"
End Function

Private Sub cmdSend_Click()

    mIntLib.State = "Q"
    mIntLib.SndPhase = 0
    mIntLib.FrameNo = 0
    
    MSComm.Output = ENQ
    Call mIntLib.WriteLog(ENQ, ccPCLog)

End Sub

Private Sub Command1_Click()
strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||Bio-Rad|IH v5.0||||||||20170405140445" & vbCr
strBuffer = strBuffer & "92" & vbCrLf
strBuffer = strBuffer & "2P|1||170010535233||^|||||||||||||||||||||||||||^" & vbCr
strBuffer = strBuffer & "5D" & vbCrLf
strBuffer = strBuffer & "3O|1||170010535233^^^\^^^|^^^MO31^^^|R|20170405111036||||||||||||1||||20170405133109|||F" & vbCr
strBuffer = strBuffer & "4D" & vbCrLf
strBuffer = strBuffer & "4R|1|^^^AntiA^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|40^^|C||||R||pmc1111^|20170405111132|20170405133109|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "99" & vbCrLf
strBuffer = strBuffer & "5C|1|ID-Diluent 2^^05761.05.30^20181130\^^^|CAS^5009203031712222577^50092.03.03^20171231^1||" & vbCr
strBuffer = strBuffer & "4A" & vbCrLf
strBuffer = strBuffer & "6R|2|^^^AntiB^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|0^^|C||||R||pmc1111^|20170405111132|20170405133109|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "69" & vbCrLf
strBuffer = strBuffer & "7C|1|ID-Diluent 2^^05761.05.30^20181130\^^^|CAS^5009203031712222577^50092.03.03^20171231^2||" & vbCr
strBuffer = strBuffer & "4D" & vbCrLf
strBuffer = strBuffer & "0R|3|^^^AntiD^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|40^^|C||||R||pmc1111^|20170405111132|20170405133109|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "9A" & vbCrLf
strBuffer = strBuffer & "1C|1|ID-Diluent 2^^05761.05.30^20181130\^^^|CAS^5009203031712222577^50092.03.03^20171231^3||" & vbCr
strBuffer = strBuffer & "48" & vbCrLf
strBuffer = strBuffer & "2R|4|^^^CtrlAB^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|0^^|C||||R||pmc1111^|20170405111132|20170405133109|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "B1" & vbCrLf
strBuffer = strBuffer & "3C|1|ID-Diluent 2^^05761.05.30^20181130\^^^|CAS^5009203031712222577^50092.03.03^20171231^4||" & vbCr
strBuffer = strBuffer & "4B" & vbCrLf
strBuffer = strBuffer & "4R|5|^^^cellA1^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|0^^|C||||R||pmc1111^|20170405111132|20170405133109|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "AE" & vbCrLf
strBuffer = strBuffer & "5C|1|ID-DiaCell A1^^06012.79.1^20170424\^^^|CAS^5009203031712222577^50092.03.03^20171231^5||" & vbCr
strBuffer = strBuffer & "1A" & vbCrLf
strBuffer = strBuffer & "6R|6|^^^cellB^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|30^^|C||||R||pmc1111^|20170405111132|20170405133109|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "B4" & vbCrLf
strBuffer = strBuffer & "7C|1|ID-DiaCell B^^06032.79.1^20170424\^^^|CAS^5009203031712222577^50092.03.03^20171231^6||" & vbCr
strBuffer = strBuffer & "EF" & vbCrLf
strBuffer = strBuffer & "0R|7|^^^Result^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|A^POS^^^^^^^|C||||R||pmc1111^pmc1111|20170405111132|20170405133109|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "54" & vbCrLf
strBuffer = strBuffer & "1C|1|^^^||" & vbCr
strBuffer = strBuffer & "BF" & vbCrLf
strBuffer = strBuffer & "2O|2||170010534410^^^\^^^|^^^PR15A^^^|R|20170405111037||||||||||||1||||20170405132709|||F" & vbCr
strBuffer = strBuffer & "98" & vbCrLf
strBuffer = strBuffer & "3R|1|^^^AbsI1^PR15A^Ab.screening: I,II (IAT) (5053)^|0^^|C||||R||pmc1111^|20170405113555|20170405132709|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "6C" & vbCrLf
strBuffer = strBuffer & "4C|1|ID-DiaCell I^^06111.76.1^20170424\^^^|CAS^5053104101705473273^50531.04.10^20170531^1||" & vbCr
strBuffer = strBuffer & "E8" & vbCrLf
strBuffer = strBuffer & "5R|2|^^^AbsI2^PR15A^Ab.screening: I,II (IAT) (5053)^|0^^|C||||R||pmc1111^|20170405113555|20170405132709|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "70" & vbCrLf
strBuffer = strBuffer & "6C|1|ID-DiaCell II^^06121.76.1^20170424\^^^|CAS^5053104101705473273^50531.04.10^20170531^2||" & vbCr
strBuffer = strBuffer & "35" & vbCrLf
strBuffer = strBuffer & "7R|3|^^^Result^PR15A^Ab.screening: I,II (IAT) (5053)^|^^NEG^^^^^^|C||||R||pmc1111^pmc1111|20170405113555|20170405132709|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "43" & vbCrLf
strBuffer = strBuffer & "0C|1|^^^||" & vbCr
strBuffer = strBuffer & "BE" & vbCrLf
strBuffer = strBuffer & "1L|1|N" & vbCr
strBuffer = strBuffer & "04" & vbCrLf
strBuffer = strBuffer & ""

strBuffer = ""
strBuffer = strBuffer & "1H|\^&|||Bio-Rad|IH v5.0||||||||20170425102903" & vbCr
strBuffer = strBuffer & "91" & vbCrLf
strBuffer = strBuffer & "2P|1|05-20170425-23|05-20170425-23||1^|||U||||||||||||||||||||||||^" & vbCr
strBuffer = strBuffer & "F7" & vbCrLf
strBuffer = strBuffer & "3O|1||170011461487^^^\^^^|^^^MO31^^^|R|20170425095250||||||||||||1||||20170425102903|||F" & vbCr
strBuffer = strBuffer & "62" & vbCrLf
strBuffer = strBuffer & "4R|1|^^^AntiA^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|0^^|C||||R||pmc1111^|20170425095553|20170425102903|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "79" & vbCrLf
strBuffer = strBuffer & "5C|1|ID-Diluent 2^^05761.05.30^20181130\^^^|CAS^5009203031712222600^50092.03.03^20171231^1||" & vbCr
strBuffer = strBuffer & "3D" & vbCrLf
strBuffer = strBuffer & "6R|2|^^^AntiB^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|0^^|C||||R||pmc1111^|20170425095553|20170425102903|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "7D" & vbCrLf
strBuffer = strBuffer & "7C|1|ID-Diluent 2^^05761.05.30^20181130\^^^|CAS^5009203031712222600^50092.03.03^20171231^2||" & vbCr
strBuffer = strBuffer & "40" & vbCrLf
strBuffer = strBuffer & "0R|3|^^^AntiD^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|40^^|C||||R||pmc1111^|20170425095553|20170425102903|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "AE" & vbCrLf
strBuffer = strBuffer & "1C|1|ID-Diluent 2^^05761.05.30^20181130\^^^|CAS^5009203031712222600^50092.03.03^20171231^3||" & vbCr
strBuffer = strBuffer & "3B" & vbCrLf
strBuffer = strBuffer & "2R|4|^^^CtrlAB^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|0^^|C||||R||pmc1111^|20170425095553|20170425102903|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "C5" & vbCrLf
strBuffer = strBuffer & "3C|1|ID-Diluent 2^^05761.05.30^20181130\^^^|CAS^5009203031712222600^50092.03.03^20171231^4||" & vbCr
strBuffer = strBuffer & "3E" & vbCrLf
strBuffer = strBuffer & "4R|5|^^^cellA1^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|40^^|C||||R||pmc1111^|20170425095553|20170425102903|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "F6" & vbCrLf
strBuffer = strBuffer & "5C|1|ID-DiaCell A1^^06012.81.1^20170522\^^^|CAS^5009203031712222600^50092.03.03^20171231^5||" & vbCr
strBuffer = strBuffer & "05" & vbCrLf
strBuffer = strBuffer & "6R|6|^^^cellB^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|40^^|C||||R||pmc1111^|20170425095553|20170425102903|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "C9" & vbCrLf
strBuffer = strBuffer & "7C|1|ID-DiaCell B^^06032.81.1^20170522\^^^|CAS^5009203031712222600^50092.03.03^20171231^6||" & vbCr
strBuffer = strBuffer & "DA" & vbCrLf
strBuffer = strBuffer & "0R|7|^^^Result^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|0^POS^^^^^^^|C||||R||pmc1111^automatically accept|20170425095553|20170425102903|1|IH-500|0500561|pmc1111" & vbCr
strBuffer = strBuffer & "5C" & vbCrLf
strBuffer = strBuffer & "1C|1|^^^||" & vbCr
strBuffer = strBuffer & "BF" & vbCrLf
strBuffer = strBuffer & "2L|1|N" & vbCr
strBuffer = strBuffer & "05" & vbCrLf
strBuffer = strBuffer & ""

'                                                                                  숫자임
'strBuffer = strBuffer & "0R|7|^^^Result^MO31^Bloodgr.+Rev.gr.: A,B,DVI-,ctl/A1,B (DiaClon) (5009)^|0^POS^^^^^^^|C||||R||pmc1111^pmc1111|20170405115412|20170405132751|1|IH-500|0500561|pmc1111" & vbCr
'strBuffer = strBuffer & "4A"& vbCrLf
'strBuffer = strBuffer & "1C|1|^^^||" & vbCr
'strBuffer = strBuffer & "BF"& vbCrLf
'strBuffer = strBuffer & "2L|1|N" & vbCr
'strBuffer = strBuffer & "05"& vbCrLf
'strBuffer = strBuffer & ""

'strBuffer = strBuffer & "2L|1|N" & vbCr
'strBuffer = strBuffer & "05" & vbCrLf
'strBuffer = strBuffer & ""

'strBuffer = ACK

Call MSComm_OnComm

End Sub

Private Sub Form_Activate()
    MainFrm.lblMenuNm = Me.Caption
'    Me.MDIActiveX.WindowState = ccMaximize
End Sub

'   Access DB Connect
Public Function Set_DbConnect_Jet() As Boolean
    Dim DB_Name         As String
    Dim UserName        As String
    Dim Password        As String
    Dim blnWinNTAuth    As Boolean
    Dim strSrcfile  As String
    Dim strDestFile As String

    Set AdoCn = New ADODB.Connection

On Error GoTo ConnectError


'    FileLiaison.Path = "C:\RAPID\EXPORT\"

'    DB_Name = "C:\Program Files\LG Life Sciences\AdvanSure AlloScan 2\DB\DBresults.mdb"   '   MDB Full Path & File Name
'    UserName = "admin"  '   MDB User Name (Default = 'admin')
'    Password = "reader_admin"  '   MDB Pass Word
'
'    If (DB_Name = "") Or (UserName = "") Then
'        Set_DbConnect_Jet = False
'        Set AdoCn = Nothing
'        Exit Function
'    End If
'
'    With AdoCn
'        .ConnectionTimeout = 25
'        .CursorLocation = adUseClient
'        .Provider = "Microsoft.Jet.OLEDB.4.0"
'        .Properties("Mode").Value = adModeReadWrite
'        .Properties("Persist Security Info").Value = False
'        .Properties("Data Source").Value = DB_Name
'        .Properties("User ID").Value = UserName
'        .Properties("Jet OLEDB:Database Password").Value = Password
'        .Properties("Jet OLEDB:Compact Without Replica Repair").Value = True
'        .Open
'    End With
'
'    Set_DbConnect_Jet = True
    
 Exit Function

ConnectError:
    '   오류처리
    MsgBox "   Error No. : " & Err.Number & vbCrLf & _
           " Description : " & Err.Description & vbCrLf & _
           "      Source : " & Err.Source & vbCrLf & vbCrLf _
           , vbCritical, " DB Open Error"

    If AdoCn.State <> adStateOpen Then
        Set_DbConnect_Jet = False
        Set AdoCn = Nothing
    End If

End Function



Private Sub Form_Load()
    Me.Caption = mEqpKey
    
    Me.MousePointer = vbHourglass
    
    Set mIntErrors = New clsIISIntErrors
    Set mIntLib = New clsIISInterface
    Set mOrder = New clsIISIntOrder
    
    Call CtlClear
    Call mIntLib.SetConfig(mEqpCd, mEqpKey)
    Call GetEqpComm
    DoEvents
    
    dtpFrDate.Value = Now - 1
    dtpToDate.Value = Now
    
    mIntLib.Phase = 1               '전송상태 초기화
    
    chkTmr.Value = 1
    txtTmrSec.Text = 30
    
    'tmrResult.Interval = 60000
    'tmrResult.Interval = 30000
'    tmrResult.Interval = 3000
    'tmrResult.Enabled = True
    
    intPatSeq = 0
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Deactivate()
'    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISLiaison = Nothing
End Sub



Private Sub cmdAlarm_Click()
    Dim objErrorShow As clsIISErrorShow     '에러폼 표시 클래스
    
    Set objErrorShow = New clsIISErrorShow
    Call objErrorShow.ShowErrors(mIntErrors)
    Set objErrorShow = Nothing
    
    '## 에러가 없으면 버튼색깔 원래대로, 있으면 계속 빨강색
    cmdAlarm.BackColor = IIf(mIntErrors.Count = 0, &HF4F0F2, vbRed)
    
    '## 1.0.1: 이상대(2005-02-22)
    '   - Alarm창이 닫힌후 포커스를 txtBarNo로 이동
'    txtBarNo.SetFocus

End Sub

Private Sub cmdClear_Click()
    Call CtlClear
    Call mIntLib.AccInfos.RemoveAll
    
'    txtBarNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    
'    With MSComm
'        '## 이미 포트가 열린경우
'        If .PortOpen Then
'            If MsgBox(Me.Caption & " 장비와 연결되어 있습니다." & vbNewLine & vbNewLine & _
'                      Me.Caption & " 인터페이스를 종료하시겠습니까?", vbYesNo + vbCritical) = vbYes Then
'                Unload Me
'            End If
'        End If
'
'    End With
    
    Unload Me
End Sub

Private Sub tmrResult_Timer()
    Dim intRow As Integer
    
    'Call cmdClear_Click
    
    Call cmdSearch_Click
    
    With tblReady
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = TReadyEnum.ccNo
            If Trim(.Text) <> "SEND" Then
                Call cmdSend_Click
                Exit For
            End If
        Next
    End With
    
End Sub

Private Sub Get_SearchList()
    Dim intRow      As Long
    Dim strAge      As String
    Dim strTransDt  As String
    Dim strHMsg     As String
    Dim strDMsg     As String
    
    Dim objIntInfo   As clsIISIntInfo    '인터페이스 검체정보 클래스
    Dim objIntNms    As clsIISIntNms     '장비별 검사항목 컬렉션 클래스
    Dim objBuffer    As clsIISBuffer     '버퍼 클래스
    
    Dim vWorkNo      As Variant  'Spread의 WorkNo
    Dim vBarNo       As Variant  'Spread의 바코드번호
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarno     As String   '수신한 BarNO
    Dim strWorkNo    As String   '수신한 WorkNo
    Dim strIntResult As String   '수신한 검사결과
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   'LIS 검사결과
    Dim strClass     As String   'Class 판정결과
    Dim strTemp      As String
    Dim i            As Long
'    Dim intRow       As Integer
    Dim blnSameBar   As Boolean
    Dim intCnt As Integer
    
    Dim y, Y1, Y2, Y3, X1, X2
    
    Set objIntNms = mIntLib.IntNms
    

    Set AdoRS = Get_ResultList

    If Not AdoRS.BOF Then
'        tblComplete.MaxRows = AdoRS.RecordCount
        intRow = 1: strTransDt = ""
        Do Until AdoRS.EOF
'           Call EditRcvData
            strBarno = AdoRS.Fields("PATIENTID").Value
            Set objIntInfo = New clsIISIntInfo
            With objIntInfo
                .BarNo = strBarno
                .SpcPos = strWorkNo
            End With

            For intCnt = 1 To 42
                If intCnt = 1 Then
                    If UCase(Trim(AdoRS.Fields("STRIPPANEL_A").Value & "")) = "FOOD" Then
                        strIntBase = "FD"
                    ElseIf UCase(Trim(AdoRS.Fields("STRIPPANEL_A").Value & "")) = "INHALANT" Then
                        strIntBase = "IN"
                    End If
                ElseIf intCnt = 22 Then
                    If UCase(Trim(AdoRS.Fields("STRIPPANEL_B").Value & "")) = "FOOD" Then
                        strIntBase = "FD"
                    ElseIf UCase(Trim(AdoRS.Fields("STRIPPANEL_A").Value & "")) = "INHALANT" Then
                        strIntBase = "IN"
                    End If
                End If
                
                
                strIntBase = Mid(strIntBase, 1, 2) & Format(intCnt, "00")
                
                If intCnt <= 21 Then
                    strResult = "BANDVAL_A" & intCnt + 1
                Else
                    strResult = "BANDVAL_B" & intCnt - 20
                End If
                
                strIntResult = AdoRS.Fields(strResult).Value
                
'                On Error Resume Next
                If Mid(strIntBase, 1, 2) = "FD" Then
                    Select Case intCnt
                    Case 9, 16, 22, 26, 27, 30, 37, 40 '-- 함수A
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        If Y2 >= 0 Then
                            Y3 = Log(Y2)
                        Else
                            Y3 = Log(Abs(Y2))
                        End If
                        X1 = (Y3 - AN) / AM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 1, 3, 4, 6, 7, 11, 12, 15, 18, 20, 24, 39 '-- 함수B
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        'Y3 = Log(Y2)
                        If Y2 >= 0 Then
                            Y3 = Log(Y2)
                        Else
                            Y3 = Log(Abs(Y2))
                        End If
                        X1 = (Y3 - BN) / BM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 2, 5, 8, 10, 13, 14, 17, 19, 21, 23, 25, 28, 29, 31, 32, 33, 34, 35, 36, 38, 41, 42 '-- 함수C
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        'Y3 = Log(Y2)
                        If Y2 >= 0 Then
                            Y3 = Log(Y2)
                        Else
                            Y3 = Log(Abs(Y2))
                        End If
                        
                        X1 = (Y3 - Cn) / CM
                        X2 = Exp(X1)
                        strIntResult = X2
                    End Select
                    

                ElseIf Mid(strIntBase, 1, 2) = "IN" Then
                    Select Case intCnt
                    Case 13, 18, 21, 22, 24, 25, 26, 29, 37, 39, 40 '-- 함수A
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        Y3 = Log(Abs(Y2))
                        X1 = (Y3 - AN) / AM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 1, 3, 5, 6, 7, 9, 12, 14, 17, 19, 30, 38 '-- 함수B
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        Y3 = Log(Abs(Y2))
                        X1 = (Y3 - BN) / BM
                        X2 = Exp(X1)
                        strIntResult = X2
                    Case 2, 4, 8, 10, 11, 15, 16, 20, 23, 27, 28, 31, 32, 33, 34, 35, 36, 41, 42 '-- 함수C
                        y = strIntResult
                        Y1 = (y - VarD) / (VarA - VarD)
                        Y2 = Y1 / (1 - Y1)
                        Y3 = Log(Abs(Y2))
                        X1 = (Y3 - Cn) / CM
                        X2 = Exp(X1)
                        strIntResult = X2
                    End Select
                End If

                If strIntBase = "FD01" Or strIntBase = "IN01" Then
                    If strIntResult > 100 Then
                        strIntResult = ">100"
                        strClass = "P 증가" '증가됨
                    ElseIf strIntResult <= 100 Then
                        strIntResult = "=<100"
                        strClass = "N 정상" '정상
                    End If
                Else
                    If strIntResult < 0.35 Then
                        strClass = "0"
                        '-- 2010.04.07 단대 성백달 선생님 요구사항
                        '-- Class 0 은 "0.00" 으로 치환한다.
                        'strIntResult = Format(strIntResult, "#0.#0")
                        strIntResult = "0.00"
                    ElseIf strIntResult >= 0.35 And strIntResult < 0.7 Then
                        strClass = "1" '"*"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 0.7 And strIntResult < 3.5 Then
                        strClass = "2" '"**"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 3.5 And strIntResult < 17.5 Then
                        strClass = "3" '"***"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 17.5 And strIntResult < 50 Then
                        strClass = "4" '"****"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 50 And strIntResult < 100 Then
                        strClass = "5" '"*****"
                        strIntResult = Format(strIntResult, "#0.#0")
                    ElseIf strIntResult >= 100 Then
                        strClass = "6" '"******"
                        strIntResult = ">=100"
                    Else
                        strIntResult = "0.00"
                        strClass = ""
                    End If
                End If
                If strIntResult = "0." Then strIntResult = "0.00"

                strIntResult = strIntResult & "  " & strClass
                strResult = strIntResult
'                If strIntBase = "FD01" Or strIntBase = "IN01" Then
'                '    strIntResult = strIntResult & " " & strClass
'                Else
'                    strIntResult = strIntResult & " " & strClass
'                End If
'                x >= 100
'                50 =< x <100
'                17.5 =< x < 50
'                3.5 =< x < 17.5
'                0.7 =< x < 3.5
'                0.35 =< x < 0.7
'                x < 0.35

                '## 결과값에 "?", ">", "<" 포함되어 있으면 에러로 표시
                'If IsNumeric(strIntResult) Then
                '    strResult = strIntResult
                'Else
                '    strResult = IISERROR
                'End If

                If objIntNms.ExistIntBase(strIntBase) Then
                    Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                         strIntResult, strResult, strClass)
                End If
            Next
            
            Call SaveServer(objIntInfo)
            
            intRow = intRow + 1
            AdoRS.MoveNext
        Loop
    End If
    
    Set AdoRS = Nothing
    tmrResult.Enabled = False
End Sub

'   Result List Recordset
Public Function Get_ResultList() As ADODB.Recordset
    Dim strSql      As String

On Error GoTo ErrorTrap

             strSql = "SELECT * "
    strSql = strSql & "  FROM RESULTS "
'    strSql = strSql & " WHERE EXAMDATE = '" & Format(Now, "yyyy-mm-01") & "' "
'    strSql = strSql & " WHERE PATIENTID = '10000189895' "
    
    Set AdoRS = New ADODB.Recordset
    If Get_Recordset(AdoCn, strSql, AdoRS, "") Then
        Set Get_ResultList = AdoRS
        'blnRS = True
    Else
        Set Get_ResultList = Nothing
        'blnRS = False
    End If
    
    Set AdoRS = Nothing

Exit Function

ErrorTrap:
    Set AdoRS = Nothing
'    blnRS = False

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

Private Sub txtBarNo_GotFocus()
    With txtBarNo
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Public Sub txtBarNo_KeyDown(KeyCode As Integer, Shift As Integer)
    '## 해당 바코드번호에 대한 오더정보 조회
    If KeyCode = vbKeyReturn Then
        Me.MousePointer = vbHourglass
        Call GetOrder(Trim(txtBarNo.Text))
        Me.MousePointer = vbDefault
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
'''    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
'''    Dim vBarNo      As Variant          'Spread의 바코드번호
'''    Dim strSpcYy    As String           '검체연도
'''    Dim lngSpcNo    As Long             '검체번호
'''
'''    If Row = 0 Then Exit Sub
'''
''''    Call CtlClear(ccLabel)
''''    Call mTblClear(tblResult)
'''    Call tblReady.GetText(TReadyEnum.ccBarNo, Row, vBarNo)
'''    If vBarNo = "" Then Exit Sub
'''
''''    Call GetOrder(vBarNo)
'''
'''    Set objAccInfo = mIntLib.GetAccInfo(vBarNo)
'''    If Not (objAccInfo Is Nothing) Then
'''        '## tblReady, tblresult, Label에 정보표시
''''        Call SetReady(objAccInfo)
'''        Call SetLabel(objAccInfo)
'''        Call SetResult(objAccInfo)
'''
'''
''''        Call SetOrderWS(objAccInfo)
'''
'''        Set objAccInfo = Nothing
'''    End If
'''    txtBarNo.Text = "": txtBarNo.SetFocus
'''
''''    strSpcYy = Mid$(vBarNo, 1, SPCYYLEN)
''''    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
''''    Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
'''
'''
'''    '## tblResult, Label에 정보표시
''''    Call SetLabel(objAccInfo)
''''    Call SetResult(objAccInfo)
'''
'''    Set objAccInfo = Nothing

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
    Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
    
    '## tblResult, Label에 정보표시
    Call SetLabel(objAccInfo)
    Call SetResult(objAccInfo)
    
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
    Dim strResult   As String   'LIS 결과
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
            '-- 추가
            '.Col = TCompleteEnum.ccNo:      lblPnlNm.Caption = IIf(.Text = "IN", "INHALANT", "FOOD")

        ElseIf strQcFg = "1" Then
            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
        End If
        
        For i = TCompleteEnum.ccResult To .DataColCnt
            .Col = i:   strTemp = .Text
            
            '## 1.0.3: 이상대(2005-06-24)
            '   - 화면표시 버그수정
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
            tblResult.Col = TResultEnum.ccEqpResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
            tblResult.Col = TResultEnum.ccClass:        tblResult.Text = mGetP(strTemp, TResultEnum.ccClass, DIV)
            tblResult.Col = TResultEnum.ccLISResult
                strResult = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
                tblResult.Text = strResult
                If strResult = IISERROR Then
                    tblResult.ForeColor = vbRed
                Else
                    tblResult.ForeColor = vbBlack
                End If
        Next i
        
        'Call AllergyPrint(Col, Row)
        
    End With

End Sub



'Private Sub AllergyPrint(ByVal sPtid As Long, ByVal sPanel As Long)
'    Dim strPanel    As String
'    Dim strBarNo    As String
'    Dim strResult   As String   'LIS 결과
'    Dim strTemp     As String
'    Dim i           As Long
'    Dim j           As Integer
'
'
'    'Call mTblClear(vasPrint)
'    With tblComplete
'        .Row = Row
'        .Col = TCompleteEnum.ccNo:    strPanel = .Text
'        If strPanel = "IN" Then
'            strPanel = "INHALANT"
'        ElseIf strPanel = "FD" Then
'            strPanel = "FOOD"
'        End If
'        .Col = TCompleteEnum.ccBarNo:    strBarNo = .Text
'        If strBarNo <> "" Then
'            .Col = TCompleteEnum.ccPtId:    lblPtId.Caption = .Text
'            .Col = TCompleteEnum.ccName:    lblName.Caption = .Text
'            .Col = TCompleteEnum.ccSexAge:  lblSexAge.Caption = .Text
'            .Col = TCompleteEnum.ccDoctNm:  lblDoctNm.Caption = .Text
'            .Col = TCompleteEnum.ccDeptNm:  lblDeptNm.Caption = .Text
'            .Col = TCompleteEnum.ccWardNm:  lblWardNm.Caption = .Text
'            .Col = TCompleteEnum.ccStatFg:  lblStatFg.Caption = .Text
'            .Col = TCompleteEnum.ccSpcNm:   lblSpcNm.Caption = .Text
'        End If
'
'        '-- 알러지 Information
'        Call vasPrint.SetText(1, 4, "ID:")
'        Call vasPrint.SetText(2, 4, strBarNo)
'        Call vasPrint.SetText(3, 4, "AGE:")
'        Call vasPrint.SetText(4, 4, mGetP(lblSexAge.Caption, 2, "/"))
'        Call vasPrint.SetText(5, 4, "의뢰과:")
'        Call vasPrint.SetText(6, 4, lblWardNm.Caption)
'
'        Call vasPrint.SetText(1, 5, "성명:")
'        Call vasPrint.SetText(2, 5, lblName.Caption)
'        Call vasPrint.SetText(3, 5, "SEX:")
'        Call vasPrint.SetText(4, 5, IIf(mGetP(lblSexAge.Caption, 1, "/") = "M", "남자", "여자"))
'        Call vasPrint.SetText(5, 5, "검사자:")
'        Call vasPrint.SetText(6, 5, lblDoctNm.Caption)
'
'        Call vasPrint.SetText(1, 6, "검사일:")
'        Call vasPrint.SetText(2, 6, lblName.Caption)
'        Call vasPrint.SetText(3, 6, "PANEL:")
'        Call vasPrint.SetText(4, 6, strPanel)
'        Call vasPrint.SetText(5, 6, "확인자:")
'        Call vasPrint.SetText(6, 6, lblDoctNm.Caption)
'
'        '-- 알러지 Panel
'        If strPanel = "INHALANT" Then
'            Call vasPrint.SetText(3, 10, "INHALANT 1KO PANEL")
'        ElseIf strPanel = "FOOD" Then
'            Call vasPrint.SetText(3, 10, "FOOD 3KO PANEL")
'        End If
'
'        '-- 알러지 결과
'        Call vasPrint.SetText(2, 12, "Control")
'        Call vasPrint.SetText(3, 12, "[POSITIVE]")
'        j = 1
'        For i = TCompleteEnum.ccResult To .DataColCnt
'            .Col = i:   strTemp = .Text
'            Call vasPrint.SetText(1, j + 13, CStr(j))
'            Call vasPrint.SetText(2, j + 13, mGetP(strTemp, TResultEnum.ccTestNm, DIV))
'            Call vasPrint.SetText(3, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            Call vasPrint.SetText(4, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            Call vasPrint.SetText(5, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            Call vasPrint.SetText(6, j + 13, mGetP(strTemp, TResultEnum.ccEqpResult, DIV))
'            j = j + 1
'            'vasPrint.Col = TResultEnum.ccTestNm:       vasPrint.Text = mGetP(strTemp, TResultEnum.ccTestNm, DIV)
'            'vasPrint.Col = TResultEnum.ccUnit:         vasPrint.Text = mGetP(strTemp, TResultEnum.ccUnit, DIV)
'            'vasPrint.Col = TResultEnum.ccHLDiv:        vasPrint.Text = mGetP(strTemp, TResultEnum.ccHLDiv, DIV)
'            'vasPrint.Col = TResultEnum.ccDPDiv:        vasPrint.Text = mGetP(strTemp, TResultEnum.ccDPDiv, DIV)
'            'vasPrint.Col = TResultEnum.ccRef:          vasPrint.Text = mGetP(strTemp, TResultEnum.ccRef, DIV)
'            'vasPrint.Col = TResultEnum.ccEqpResult:    vasPrint.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
'            'vasPrint.Col = TResultEnum.ccClass:        vasPrint.Text = mGetP(strTemp, TResultEnum.ccClass, DIV)
'            'vasPrint.Col = TResultEnum.ccLISResult
'        Next i
'    End With
'
'End Sub


'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '송신할 데이터
    Dim intRow    As Integer
    Dim strOrder  As String
    Dim strBarno  As String
    Dim strPtID   As String
    
    With tblReady
        For intRow = 1 To .DataRowCnt
            .Row = intRow
            .Col = TReadyEnum.ccNo
            If Trim(.Text) <> "SEND" Then
                Select Case mIntLib.SndPhase
                    Case 0  '## Header
                                                       '1H|\^&|||Liaison|||||LaborEDV|||1|20180116111106
                                                       '1H|\^&|||Liaison|||||LaborEDV|||1|20180116111106
                                                         'H|\^&|||Liaison|||||LaborEDV|||1|19941115202738
'                                                        1H|\^&|||Liaison|||||GCRL_IF |||1|20110216103538¶8C¶
                       ' strOutput = mIntLib.GetFrameNo & "H|\^&|||Liaison|||||LaborEDV|||1|" & Format(Now, "yyyymmddhhmmss") & vbCr & ETX
                        strOutput = mIntLib.GetFrameNo & "H|\^&|||" & vbCr & ETX
                        mIntLib.SndPhase = 1
                        intPatSeq = 0
                        
                    Case 1  '## Patient
                        .Row = intRow
                        .Col = TReadyEnum.ccPtId
                        strPtID = Trim(.Text)
                        intPatSeq = intPatSeq + 1
                                                       '2P|1||406221230372   ||^|||||||||||||||||||||¶19¶
                        strOutput = mIntLib.GetFrameNo & "P|" & CStr(intPatSeq) & "||" & strPtID & "||^|||||||||||||||||||||" & vbCr & ETX
                        mIntLib.SndPhase = 2
                        
                    Case 2  '## Order
                        .Row = intRow
                        .Col = TReadyEnum.ccBarNo
                        strBarno = Trim(.Text)
                        .Col = TReadyEnum.ccOrder
                        strOrder = Trim(.Text)
                        
                                                       '3O|1|406221230372    ||^^^25OHD^       |N|20110215                       |||||||||||||||||||O|¶9E¶
                        strOutput = mIntLib.GetFrameNo & "O|1|" & strBarno & "||" & strOrder & "|N|" & Format(Now, "yyyymmdd") & "|||||||||||||||||||O|" & vbCr & ETX
                        
                        'mOrder.GetOrder
                        
                        If intRow <> .DataRowCnt Then
                            .SetText TReadyEnum.ccNo, intRow, "SEND"
                        End If
                        
                        If intRow >= .DataRowCnt Then
                            mIntLib.SndPhase = 3
                        Else
                            mIntLib.SndPhase = 1
                        End If
                    
                    Case 3  '## Termianator
                        strOutput = mIntLib.GetFrameNo & "L|1|N" & vbCr & ETX
                        mIntLib.SndPhase = 4
                    Case 4  '## EOT
                        mIntLib.State = ""
                        MSComm.Output = EOT
                        Call mIntLib.WriteLog(EOT, ccPCLog)
                        .SetText TReadyEnum.ccNo, intRow, "SEND"
                        Exit Sub
                End Select
                    
                strOutput = STX & strOutput & fGetChkSum(strOutput) & vbCrLf
                MSComm.Output = strOutput
                Debug.Print strOutput
                Call mIntLib.WriteLog(strOutput, ccPCLog)
                
                'If mIntLib.SndPhase = 3 Then
                '    .SetText TReadyEnum.ccNo, intRow, "SEND"
                'End If
                Exit For
            End If
        Next
    End With
    
End Sub

Public Function fGetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        fGetChkSum = "00"
    Else
        fGetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
    
End Function

Private Sub MSComm_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long


'GoTo rst
    Select Case MSComm.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long
            
            Buffer = MSComm.Input
'rst:
           ' Buffer = strBuffer
            Call mIntLib.WriteLog(Buffer, ccEqp)

            lngBufLen = Len(Buffer)
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case mIntLib.Phase
                    Case 1      '## Estabilshment Phase
                        Select Case BufChar
                            Case ENQ
                                mIntLib.BufCnt = 1
                                Call mIntLib.ClearBuffer
                                mIntLib.Phase = 2
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                            Case ACK
                                If mIntLib.State = "Q" Then
                                    Call SendOrder
                                End If
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                mIntLib.BufCnt = 1
                                Call mIntLib.ClearBuffer
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                            Case STX
                            Case vbCr
                                mIntLib.BufCnt = mIntLib.BufCnt + 1
                            Case ETX
                                mIntLib.Phase = 3
                            Case ETB
                                mIntLib.Phase = 3
                                mIntLib.IsETB = True
                            Case Else
                                If mIntLib.IsETB = False Then
                                    Call mIntLib.AddBuffer(BufChar)
                                Else
                                    mIntLib.IsETB = False
                                End If
                        End Select
                    Case 3      '## Transfer Phase
                        Select Case BufChar
                            Case vbCr
                            Case vbLf
                                mIntLib.Phase = IIf(mIntLib.IsETB = False, 4, 2)
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                        End Select
                    Case 4      '## Termination Phase
                        Select Case BufChar
                            Case STX
                                mIntLib.Phase = 2
                            Case EOT
                                Call EditRcvData
                                If mIntLib.State = "Q" Then
                                    mIntLib.SndPhase = 0
                                    mIntLib.FrameNo = 0
                                    MSComm.Output = ENQ
                                    Call mIntLib.WriteLog(ENQ, ccPCLog)
                                End If
                                mIntLib.Phase = 1
                        End Select
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
'   기능 : 장비로부 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim objIntInfo   As clsIISIntInfo    '인터페이스 검체정보 클래스
    Dim objIntNms    As clsIISIntNms     '장비별 검사항목 컬렉션 클래스
    Dim objBuffer    As clsIISBuffer     '버퍼클래스

    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarno     As String   '수신한 바코드번호
    Dim strRackNo    As String   '수신한 Rack No
    Dim strPos       As String   '수신한 Sample Position
    Dim strIntPart   As String   '수신한 ABO/AbScreening
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과
    Dim strIntResult    As String   '수신한 결과
    Dim strTemp      As String
    Dim strNRQC      As String
    
    Set objIntNms = mIntLib.IntNms
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers

        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
                mIntLib.State = ""
                
            Case "Q"    '## Request Information
                '2Q|1|404811201912||ALL||||||||O¶1D¶
                '## 바코드번호, Sample Type, Rack No, Position, Kind, Priority 조회
                strBarno = mGetP(strRcvBuf, 3, "|")

                With mOrder
                    .ClsClear
                    .BarNo = strBarno
                End With
                Call GetOrder(strBarno)
                mIntLib.State = "Q"

            Case "P"    '## Patient
            Case "O"    '## Order
'3O|1|180008322875||^^^Myco-M^|||||||||||||||||||||F
'1B
                '## DB에 결과저장
                If mIntLib.State = "R" And strBarno <> "" Then
                    Call SaveServer(objIntInfo)
                    Set objIntInfo = Nothing
                    mIntLib.State = ""
                End If
            
                '## 바코드번호, Rack No, Sample Position 조회
                strTemp = mGetP(strRcvBuf, 3, "|")
                strBarno = mGetP(strTemp, 1, "^")
                'strBarno = Mid(strBarno, 1, 11)
                
                Set objIntInfo = New clsIISIntInfo
                With objIntInfo
                    .BarNo = strBarno
                End With
            
                
            Case "R"    '## Result
'4R | 1 | ^^^Myco-M |2.1 |Index||N||F||||20180115123803|Liaison
'EE
                strTemp = mGetP(strRcvBuf, 3, "|")
                strIntBase = Trim$(mGetP(strTemp, 4, "^"))
                strIntResult = mGetP(strRcvBuf, 4, "|")
                strResult = mGetP(strRcvBuf, 7, "|")
                                
                strIntResult = Replace(strIntResult, "<", "")
                strIntResult = Replace(strIntResult, ">", "")
        
                If strIntBase = "Myco-M" Then
                    If strIntResult >= 27 Then
                        strResult = "POSITIVE:" & ">" & strIntResult
                    ElseIf strIntResult >= 10 Then
                        strResult = "POSITIVE:" & strIntResult
                    ElseIf strIntResult <= 0.1 Then
                        strResult = "NEGATIVE:" & "<" & strIntResult
                    ElseIf strIntResult < 10 Then
                        strResult = "NEGATIVE:" & strIntResult
                    Else
                        strResult = "EQUIVOCAL:" & strIntResult
                    End If
                End If
                
                If objIntNms.ExistIntBase(strIntBase) Then
                    Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                         strResult, strResult)
                End If
                mIntLib.State = "R"

            Case "L"    '## Terminator
                '## DB에 결과저장
                If mIntLib.State = "R" Then
                    Call SaveServer(objIntInfo)
                    Set objIntInfo = Nothing
                    mIntLib.State = ""
                End If
        End Select
    Next
    Set objIntNms = Nothing
    Set objBuffer = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 결과판정, 결과저장, 화면표시
'   인수 :
'       - pIntInfo : 인터페이스 검체정보 클래스
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pIntInfo As clsIISIntInfo)
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim vBarNo      As Variant 'Spread의 바코드번호
    Dim strBarno    As String  '바코드번호
    Dim strSpcYy    As String  '검체연도
    Dim lngSpcNo    As Long    '검체번호
    Dim i           As Long
    
    Me.MousePointer = vbHourglass
    
    strBarno = pIntInfo.BarNo
    
    '## 결과판정
    If mIntLib.CheckResult(pIntInfo) = -1 Then
        '## 접수정보가 없을때 결과표시
        Call SetComplete1(pIntInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Me.MousePointer = vbDefault
        Exit Sub
    Else
        '## 접수정보가 있을때 결과표시
        strSpcYy = Mid$(strBarno, 1, SPCYYLEN)
        lngSpcNo = CLng(Mid$(strBarno, SPCYYLEN + 1, SPCNOLEN))
        Set objAccInfo = mIntLib.AccInfos(strSpcYy, lngSpcNo)
        
        Call SetComplete2(objAccInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Set objAccInfo = Nothing
        
        '## ClientDb, Server에 결과저장
        Call mIntLib.SaveClientDb(strSpcYy, lngSpcNo)
        Call mIntLib.SaveResult(strSpcYy, lngSpcNo)
        Call mIntLib.Remove(strSpcYy, lngSpcNo)
        StatusBar.Panels(2).Text = "검체번호:" & strBarno & " 를 정상적으로 결과저장 했습니다."
    End If
    
    '## tblReady에서 전송된 검체삭제
    If mIntLib.BarPos = ccPC Then
        With tblReady
            For i = 1 To .DataRowCnt
                Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
                If CStr(vBarNo) = strBarno Then
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
    
    gBarNo = pBarNo
    
    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    If Not (objAccInfo Is Nothing) Then
        '## tblReady, tblresult, Label에 정보표시
        Call SetReady(objAccInfo)
        Call SetLabel(objAccInfo)
        Call SetResult(objAccInfo)
        
        
'        Call SetOrderWS(objAccInfo)
        
        Set objAccInfo = Nothing
    End If
    txtBarNo.Text = ""
    'txtBarNo.SetFocus
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Function GetInfo(ByVal pBarNo As String) As String
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    
    If pBarNo = "" Then Exit Function
    GetInfo = ""
    
    gBarNo = pBarNo
    
    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    If Not (objAccInfo Is Nothing) Then
        If objAccInfo.QcFg = "0" Then         '## 일반검체
            GetInfo = "NR"
        ElseIf objAccInfo.QcFg = "1" Then     '## QC검체
            GetInfo = "QC"
        End If
        Set objAccInfo = Nothing
    End If

End Function

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
'Private Sub GetOrderByAccNo()
'    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
'    Dim Rs   As Recordset
'    Dim vBarNo      As Variant  'Spread의 바코드번호
'    Dim strSpcYy    As String   '검체연도
'    Dim lngSpcNo    As Long     '검체번호
'
'    Set Rs = New Recordset
'    Set objAccInfo = New clsIISAccInfo
'    'rsBarcode = objAccInfo.GetTargetSpcs(mEqpCd, Format(dtpFrDate.Value, "yyyymmdd"), Format(dtpToDate.Value, "yyyymmdd"))
'
''    If pBarNo = "" Then Exit Sub
'
'    Set Rs = objAccInfo.GetTargetSpcs(mEqpCd, Format(dtpFrDate.Value, "yyyymmdd"), Format(dtpToDate.Value, "yyyymmdd"))
'    Do Until Rs.EOF
'        strSpcYy = Rs.Fields("SPCYY").Value
'        strSpcYy = Mid$(strSpcYy, 1, SPCYYLEN)
'        lngSpcNo = Rs.Fields("SPCNO").Value
'        lngSpcNo = Mid$(lngSpcNo, 1, SPCNOLEN)
'
'        vBarNo = Format$(strSpcYy, String$(SPCYYLEN, "0")) & Format$(lngSpcNo, String$(SPCNOLEN, "0"))
'
'        vBarNo = Format$(vBarNo, String$(SPCLEN, "#"))
'
'        Set objAccInfo = mIntLib.GetAccInfo(vBarNo)
'
'        If Not (objAccInfo Is Nothing) Then
'            '## tblReady, tblresult, Label에 정보표시
'            Call SetReady(objAccInfo)
'            Call SetLabel(objAccInfo)
'            Call SetResult(objAccInfo)
'
'
'            Call SetOrderWS(objAccInfo)
'
'            Set objAccInfo = Nothing
'        End If
'        Rs.MoveNext
'    Loop
'
'    Rs.Close
'    Set Rs = Nothing
''    txtBarNo.Text = "": txtBarNo.SetFocus
'
'End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblReady에 정보표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetReady(ByVal pAccInfo As clsIISAccInfo)
    Dim lngWorkNo As Long   'WorkNo
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim blnFood     As Boolean
    Dim blnIn       As Boolean
    Dim strOrder    As String
    
    blnFood = False
    blnIn = False
    
    
    With tblReady
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If
        
        '.Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
        .Col = TReadyEnum.ccBarNo:  .Text = gBarNo 'pAccInfo.GetBarNo
        .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
        
        If pAccInfo.QcFg = "0" Then         '## 일반검체
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
        ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
            .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
            .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
            '.Col = TReadyEnum.ccNo:     .Text = "ABO/SC"
        End If
        
        .Col = TReadyEnum.ccOrder:    .Text = pAccInfo.Sex & "|" & pAccInfo.Ssn
        
        strOrder = ""
        For Each objResult In pAccInfo.Results
'            If InStr(UCase(objResult.TestNm10), "ABO") > 0 Then
'                .Col = TReadyEnum.ccNo:     .Text = "ABO"
'                blnFood = True
'                Exit For
'            ElseIf InStr(UCase(objResult.TestNm10), "SCREEN") > 0 Then
'                .Col = TReadyEnum.ccNo:     .Text = "SC"
'                blnIn = True
'                Exit For
'            End If
            strOrder = strOrder & "^^^" & objResult.IntNm.IntBase & "^"
            
'            .Col = TReadyEnum.ccOrder
'            .Text = objResult.IntNm.IntBase
        
        Next
        .Col = TReadyEnum.ccOrder
        .Text = strOrder 'objResult.IntNm.IntBase
        
        '-- 다른 처방이 있는지 확인(Food면 IN을 IN이면 Food를 찾는다.
'        If blnFood = True Then
'            For Each objResult In pAccInfo.Results
'
'                If InStr(UCase(objResult.TestNm10), "SCREEN") > 0 Then
'                    If .MaxRows <= .DataRowCnt Then
'                        .MaxRows = .MaxRows + 1
'                        .Row = .MaxRows
'                    'Else
'                    '    .Row = .DataRowCnt + 1
'                    End If
'
'                    .Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
'                    .Col = TReadyEnum.ccBarNo:  .Text = gBarNo 'pAccInfo.GetBarNo
'                    .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
'
'                    If pAccInfo.QcFg = "0" Then         '## 일반검체
'                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
'                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
'                    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
'                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
'                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
'                    End If
'
'                    .Col = 6:  .Text = pAccInfo.Sex & "|" & pAccInfo.Ssn
'                    .Col = TReadyEnum.ccNo:     .Text = "ABO/SC"
'                    Exit For
'                End If
'            Next
'        End If
'
'        If blnIn = True Then
'            For Each objResult In pAccInfo.Results
'                If InStr(UCase(objResult.TestNm10), "ABO") > 0 Then
'                    If .MaxRows <= .DataRowCnt Then
'                        .MaxRows = .MaxRows + 1
'                        .Row = .MaxRows
'                    Else
'                        .Row = .DataRowCnt + 1
'                    End If
'
'                    .Col = TReadyEnum.ccNo:     .Text = CStr(lngWorkNo)
'                    .Col = TReadyEnum.ccBarNo:  .Text = gBarNo 'pAccInfo.GetBarNo
'                    .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
'
'                    If pAccInfo.QcFg = "0" Then         '## 일반검체
'                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
'                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
'                    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
'                        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.CtrlCd
'                        .Col = TReadyEnum.ccName:   .Text = pAccInfo.LevelCd
'                    End If
'
'                    .Col = 6:  .Text = pAccInfo.Sex & "|" & pAccInfo.Ssn
'                    .Col = TReadyEnum.ccNo:     .Text = "ABO"
'                    Exit For
'                End If
'            Next
'        End If

        
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
            
            '## 1.0.4: 이상대(2005-06-29)
            '   - 장비결과란에 LIS결과가 표시되는 버그수정
            .Text = objIntResult.IntNm & DIV & objIntResult.IntResult & DIV & DIV & DIV & DIV & DIV & DIV & objIntResult.Info
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
Private Sub SetComplete2(ByVal pAccInfo As clsIISAccInfo)
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

        .Col = TCompleteEnum.ccNo:      .Text = pAccInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pAccInfo.GetBarNo
        .Col = TCompleteEnum.ccAccNo:   .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
        
        If pAccInfo.QcFg = "0" Then         '## 일반검체
            .Col = TCompleteEnum.ccPtId:    .Text = pAccInfo.PtId
            .Col = TCompleteEnum.ccName:    .Text = pAccInfo.Name
            .Col = TCompleteEnum.ccSexAge:  .Text = pAccInfo.Sex & " / " & mGetAge(Mid$(pAccInfo.Ssn, 1, 6))
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
                If objResult.IntResult <> "" Then
                    .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                            DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & _
                            DIV & mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal) & DIV & DIV & objResult.IntNm.IntBase & DIV
                            '-- 2015.08.28추가 & objResult.IntNm.IntBase & DIV
                    i = i + 1
                End If
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
                        objQCResult.RstCd & DIV & objQCResult.Unit & DIV & objQCResult.RADiv & _
                        DIV & DIV & DIV
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
                .Col = TResultEnum.ccRef:       .Text = mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal)
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
'   기능 : AlloScan 워크리스트 생성 정보표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
'Private Sub SetOrderWS(ByVal pAccInfo As clsIISAccInfo)
'    Dim objResult   As clsIISResult     '결과내역 클래스
'    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
'    Dim mLogOn As clsIISLogOn
'
'    Dim strAlloFile As String
'    Dim lngFIleNum  As Long
'    Dim strInFo     As String
'    Dim strOldInFo  As String
'    Dim blnNewFlag  As Boolean
'
'    Set mLogOn = New clsIISLogOn
'
'
''    If pAccInfo.QcFg = "0" Then         '## 일반검체
'            With AlloFile
'                .CancelError = True
'                .FileName = "C:\RAPID\IMPORT\" & Trim(txtFileNm.Text) & ".asc"
''                .ShowSave
'                If Len(Dir(.FileName)) Then
'                    Kill .FileName
'                    blnNewFlag = True
'                Else
'                    blnNewFlag = False
'                End If
'
'                lngFIleNum = FreeFile
'
'                Open .FileName For Append As #lngFIleNum
'                If blnNewFlag = False Then
''[JOBLIST]
''JOBName;1/01/2009;S01-45329578234;Panel 2KO v80 UK.RDF.TST;Last01;First01;1/2/2008;MALE;
''JOBName;1/01/2009;S02-64325984359;Panel 2KO v80 UK.RDF.TST;Last01;First01;12/31/2007;MALE;
''JOBName;1/01/2009;S03-76534954562;Panel 2KO v80 UK.RDF.TST;Last03;First03;1/2/2008;MALE;
''JOBName;1/01/2009;S04-64632134814;Panel 2KO v80 UK.RDF.TST;Last04;First04;1/2/2008;MALE;
''JOBName;1/01/2009;S05-16324873488;Panel 2KO v80 UK.RDF.TST;Last05;First05;1/2/2008;MALE;
''[EOF]
'
'
'                    Print #lngFIleNum, "[JOBLIST]"
''                    Print #lngFIleNum, "JOBName;" & Format("m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";" & "Panel 2KO v80 UK.RDF.TST"
'                End If
'
'                For Each objResult In pAccInfo.Results
'                    If InStr(objResult.TestNm, "Food") > 0 Then
'                        strInFo = "1"
'                    ElseIf InStr(objResult.TestNm, "Inhalant") > 0 Then
'                        strInFo = "2"
'                    End If
'
'                    If strOldInFo <> strInFo Then
'                        If strInFo = "1" Then
'                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 1KO v80 UK.TST"
'                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 2KO v80 UK.TST"
'                        Else
'                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 3KO v80 UK.TST"
'                            Print #lngFIleNum, "JOBName;" & Format(Now, "m/dd/yyyy") & ";" & pAccInfo.GetBarNo & ";Panel 4KO v80 UK.TST"
'                        End If
'                    End If
'
'                    strOldInFo = strInFo
'                Next
'            End With
'
'        Print #lngFIleNum, "[EOF]"
'        Close #lngFIleNum
'
''    ElseIf pAccInfo.QcFg = "1" Then     '## QC검체
''        For Each objQCResult In pAccInfo.QCResults
''            With tblResult
''                If .MaxRows <= .DataRowCnt Then
''                    .MaxRows = .MaxRows + 1
''                    .Row = .MaxRows
''                Else
''                    .Row = .DataRowCnt + 1
''                End If
''
''                .Col = TResultEnum.ccTestNm:    .Text = objQCResult.IntNm.IntNm
''                .Col = TResultEnum.ccEqpResult: .Text = objQCResult.Result
''                .Col = TResultEnum.ccLISResult: .Text = objQCResult.RstCd
''                .Col = TResultEnum.ccHLDiv:     .Text = objQCResult.RADiv
''            End With
''        Next
''    End If
'
'    Set objResult = Nothing
'    Set objQCResult = Nothing
'End Sub

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
        lblSexAge.Caption = pAccInfo.Sex & " / " & mGetAge(Mid$(pAccInfo.Ssn, 1, 6))
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
            Error.SetLog App.EXEName, "frmIISABL835", "GetEqpComm", strErrMsg, Now
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
        Error.SetLog App.EXEName, "frmIISABL835", "GetEqpComm", strErrMsg, Now
        Call mIntLib_EqpError("E005")
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear(Optional ByVal pFlag As ClearEnum = ccAll)
    lblPtId.Caption = "":       lblName.Caption = ""
    lblSexAge.Caption = "":     lblDoctNm.Caption = ""
    lblDeptNm.Caption = "":     lblWardNm.Caption = ""
    lblStatFg.Caption = "":     lblSpcNm.Caption = ""
    'lblPnlNm.Caption = ""
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
                End If
            End With
        Case DELETEALL  '## Delete All
            Call mIntLib.AccInfos.RemoveAll
            Call mTblClear(tblReady)
    End Select
End Sub

'Private Sub txtFileNm_DblClick()
'    If vasINPrint.Visible = True Then
'        vasINPrint.Visible = False
'    Else
'        vasINPrint.Visible = True
'    End If
'End Sub

'Private Sub txtWorkNo_GotFocus()
'    With txtWorkNo
'        .SelStart = 0
'        .SelLength = Len(.Text)
'    End With
'End Sub

'Private Sub txtWorkNo_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        SendKeys "{TAB}"
'    End If
'End Sub
'
'Private Sub txtWorkNo_KeyPress(KeyAscii As Integer)
'    '## 숫자만 입력
'    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
'        KeyAscii = 0
'        Exit Sub
'    End If
'End Sub



VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIISLABOSPECT_2 
   BackColor       =   &H00DBE6E6&
   Caption         =   "LABOSPECT"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows 기본값
   Begin MSWinsockLib.Winsock wSck 
      Left            =   8280
      Top             =   8580
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrQ 
      Left            =   7800
      Top             =   8520
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   6548
      TabIndex        =   0
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1290
      Left            =   6548
      TabIndex        =   4
      Top             =   405
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   600
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
      Height          =   495
      Left            =   13913
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   12694
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      Height          =   495
      Left            =   11475
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8567
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6578
      Top             =   8432
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
      Height          =   3555
      Left            =   105
      TabIndex        =   21
      Top             =   525
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   6271
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
      MaxCols         =   5
      MaxRows         =   11
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISLABOSPECT_2.frx":0000
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4410
      Left            =   105
      TabIndex        =   22
      Top             =   4665
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   7779
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
      MaxCols         =   13
      MaxRows         =   14
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISLABOSPECT_2.frx":0508
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   98
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   107
      Width           =   3495
      _ExtentX        =   6165
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
      Height          =   375
      Left            =   98
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4247
      Width           =   3495
      _ExtentX        =   6165
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
      SpreadDesigner  =   "frmIISLABOSPECT_2.frx":0D58
      TextTip         =   2
   End
End
Attribute VB_Name = "frmIISLABOSPECT_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISHitachi7600.frm
'   작성자  : 오세원
'   내  용  : Hitachi 7600 장비폼
'   작성일  : 2020-11-03
'   버  전  :
'       1. 1.0.1: 오세원
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

Private WithEvents mIntLib  As clsIISInterface   '인터페이스 클래스
Attribute mIntLib.VB_VarHelpID = -1
Private WithEvents mPopup   As clsIISPopup       '팝업메뉴
Attribute mPopup.VB_VarHelpID = -1

Private mIntErrors  As clsIISIntErrors          '인터페이스 에러 컬렉션
Private mOrder      As clsIISIntOrder           '오더정보 클래스

Private mEqpCd  As String   '장비코드
Private mEqpKey As String   '장비키
Private strQState    As String
Private strSpcPos    As String

'For E-170/Hitachi7600 Serial
Dim bSTXChk     As Boolean
Dim bEndChk     As Boolean
Dim RstEnd      As String
Dim RcvBuffer   As String

Private gComType    As String
Private gSckPort    As String
Dim pBuf        As String

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property

Private Sub Command1_Click()

    pBuf = ""
    pBuf = pBuf & ""

    pBuf = ""
    pBuf = pBuf & ""
    pBuf = pBuf & "1H|\^&|||LST008AS^1|||||HOST|RSUPL^REAL|P|1" & vbCrLf
    pBuf = pBuf & "P|1|||||||U||||||^" & vbCrLf
    pBuf = pBuf & "O|1|10002229358|0^50023^4^^S1^SC|^^^989^\^^^990^\^^^991^\^^^26902^\^^^26903^\^^^28005^\^^^28010^\^^^28011^\^^^28014^\^^^28024^\^^^28056^\^^^28057^\^^^28073^\^^^28416^\^^^28421^\^93" & vbCrLf
    pBuf = pBuf & "2^^28454^\^^^28457^|R||20201110114727||||N||||1|||||||20201110115027|||F" & vbCrLf
    pBuf = pBuf & "C|1|I|                              ^^^^|G" & vbCrLf
    pBuf = pBuf & "R|1|^^^989/|      |mmol/L||A||F|||||ISE1|1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|1|I|72|I" & vbCrLf
    pBuf = pBuf & "R|2|^^^990/|      |mmol/L||A||F|||||ISE1|1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|2|I|90" & vbCrLf
    pBuf = pBuf & "372|I" & vbCrLf
    pBuf = pBuf & "R|3|^^^991/|      |mmol/L||A||F|||||ISE1|1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|3|I|72|I" & vbCrLf
    pBuf = pBuf & "R|4|^^^26902/|      |mg/dL||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|4|I|72|I" & vbCrLf
    pBuf = pBuf & "R|5|^^^26903/|      |mg/dL||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|5|I|72|I" & vbCrLf
    pBuf = pBuf & "R|6|^^^28005/|      |U/L||A||F||||||1|1||^^^|^F0" & vbCrLf
    pBuf = pBuf & "4^^" & vbCrLf
    pBuf = pBuf & "C|6|I|72|I" & vbCrLf
    pBuf = pBuf & "R|7|^^^28010/|      |mg/dL||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|7|I|72|I" & vbCrLf
    pBuf = pBuf & "R|8|^^^28011/|      |mg/dL||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|8|I|72|I" & vbCrLf
    pBuf = pBuf & "R|9|^^^28014/|      |mg/dL||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|9|I|72|I" & vbCrLf
    pBuf = pBuf & "R|10|^^^28024/|      |mg/dL||A||F||||||13B" & vbCrLf
    pBuf = pBuf & "5|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|10|I|72|I" & vbCrLf
    pBuf = pBuf & "R|11|^^^28056/|      |U/L||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|11|I|72|I" & vbCrLf
    pBuf = pBuf & "R|12|^^^28057/|      |U/L||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|12|I|72|I" & vbCrLf
    pBuf = pBuf & "R|13|^^^28073/|      |R.U.  ||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|13|I|72|I" & vbCrLf
    pBuf = pBuf & "R|14|^^^28416/|      |mg/dL5E" & vbCrLf
    pBuf = pBuf & "6||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|14|I|72|I" & vbCrLf
    pBuf = pBuf & "R|15|^^^28421/|      |mg/dL||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|15|I|72|I" & vbCrLf
    pBuf = pBuf & "R|16|^^^28454/|      |g/dL||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|16|I|72|I" & vbCrLf
    pBuf = pBuf & "R|17|^^^28457/|      |g/dL||A||F||||||1|1||^^^|^^^" & vbCrLf
    pBuf = pBuf & "C|17|I|72|I" & vbCrLf
    pBuf = pBuf & "L|1|N" & vbCrLf
    pBuf = pBuf & "BE" & vbCrLf
    pBuf = pBuf & ""

    Call MSComm_OnComm

End Sub

Private Sub Form_Activate()
    MainFrm.lblMenuNm = Me.Caption
    Me.MDIActiveX.WindowState = ccMaximize
End Sub

Private Function GetLabospectConfig(ByVal strConfigNm As String) As String
    Dim strFileName As String
    Dim strReturnedString As String
    
    GetLabospectConfig = ""
    
    strFileName = App.Path & "\Labospect.ini"
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "LABOSPECT", strConfigNm, "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    GetLabospectConfig = strReturnedString
    
End Function

Private Sub Form_Load()
    Me.Caption = mEqpKey
    
    Me.MousePointer = vbHourglass
    
    Set mIntErrors = New clsIISIntErrors
    Set mIntLib = New clsIISInterface
    Set mOrder = New clsIISIntOrder
    
    Call CtlClear
    Call mIntLib.SetConfig(mEqpCd, mEqpKey)
    
'    gComType = GetLabospectConfig("COMTYPE")    'SERIAL, TCPIP
'    gSckPort = GetLabospectConfig("TCPPORT")
'
'    If gComType = "SERIAL" Then
        Call GetEqpComm
'    Else
'        If gSckPort <> "" And IsNumeric(gSckPort) Then
'            wSck.LocalPort = CInt(gSckPort)
'            wSck.Listen
'        End If
'    End If
    
    strQState = ""
    tmrQ.Enabled = False
    
    '변수 초기화(E-170/H-7600)
    bSTXChk = False
    bEndChk = True
    RstEnd = "Y"
    RcvBuffer = ""
    
    DoEvents
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Deactivate()
    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mOrder = Nothing
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISLABOSPECT_2 = Nothing
End Sub

Private Sub cmdAlarm_Click()
    Dim objErrorShow As clsIISErrorShow     '에러폼 표시 클래스

    Set objErrorShow = New clsIISErrorShow
    Call objErrorShow.ShowErrors(mIntErrors)
    Set objErrorShow = Nothing

    '## 에러가 없으면 버튼색깔 원래대로, 있으면 계속 빨강색
    cmdAlarm.BackColor = IIf(mIntErrors.Count = 0, &HF4F0F2, vbRed)
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
    Call mIntLib.AccInfos.RemoveAll
End Sub

Private Sub cmdExit_Click()
    Unload Me
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
    Dim strIntInfo  As String   '추가정보
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

            '## 1.0.1: 이상대(2005-07-12)
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
            tblResult.Col = TResultEnum.ccLISResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
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
        
        strInfo = vbCrLf & Space(2) & strInfo & vbCrLf
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
            'Buffer = pBuf
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
                                If mIntLib.State = "Q" Then Call SendOrder
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
            
'            For i = 1 To lngBufLen
'                BufChar = Mid$(Buffer, i, 1)
'
'                Select Case mIntLib.Phase
'                    Case 1      '## Estabilshment Phase
'                        Select Case BufChar
'                            Case ENQ
'                                mIntLib.BufCnt = 1
'                                Call mIntLib.ClearBuffer
'
'                                If mIntLib.State = "Q" Then
'                                    MSComm.Output = ENQ
'                                    Call mIntLib.WriteLog(ENQ, ccPCLog)
'                                    mIntLib.State = ""
'                                    mIntLib.Phase = 1
'                                    mIntLib.SndPhase = 0
'                                    mIntLib.FrameNo = 0
'                                Else
'                                    MSComm.Output = ACK
'                                    Call mIntLib.WriteLog(ACK, ccPCLog)
'                                    mIntLib.Phase = 2
'                                End If
'                            Case ACK
'                                If strQState = "Q" Then
'                                    Call SendOrder
'                                Else
'                                    MSComm.Output = ACK
'                                    Call mIntLib.WriteLog(ACK, ccPCLog)
'                                End If
'                        End Select
'                    Case 2      '## Transfer Phase
'                        Select Case BufChar
'                            Case ENQ
'                                mIntLib.BufCnt = 1
'                                Call mIntLib.ClearBuffer
'                                MSComm.Output = ACK
'                                Call mIntLib.WriteLog(ACK, ccPCLog)
'                            Case STX
'                            Case vbCr
'                                mIntLib.BufCnt = mIntLib.BufCnt + 1
'                            Case ETX
'                                mIntLib.Phase = 3
'                            Case ETB
'                                mIntLib.Phase = 3
'                                mIntLib.IsETB = True
'                            Case Else
'                                If mIntLib.IsETB = False Then
'                                    Call mIntLib.AddBuffer(BufChar)
'                                Else
'                                    mIntLib.IsETB = False
'                                End If
'                        End Select
'                    Case 3      '## Transfer Phase
'                        Select Case BufChar
'                            Case vbCr
'                            Case vbLf
'                                mIntLib.Phase = IIf(mIntLib.IsETB = False, 4, 2)
'                                MSComm.Output = ACK
'                                Call mIntLib.WriteLog(ACK, ccPCLog)
'                        End Select
'                    Case 4      '## Termination Phase
'                        Select Case BufChar
'                            Case STX
'                                mIntLib.Phase = 2
'                            Case EOT
'                                mIntLib.Phase = 1
'                                Call EditRcvData
'
'                                tmrQ.Interval = 500
'                                tmrQ.Enabled = True
'
'                        End Select
'                End Select
'            Next i
            
'            For i = 1 To lngBufLen
'                BufChar = Mid$(Buffer, i, 1)
'
'                Select Case mIntLib.Phase
'                    Case 1
'                        Select Case BufChar
'                            Case ENQ
'                                mIntLib.Phase = 2
'                                RstEnd = "Y"
'                                bSTXChk = False
'                                bEndChk = True
'                                RcvBuffer = ""
'                                MSComm.Output = ACK
'                                Call mIntLib.WriteLog(ACK, ccPCLog)
'                            Case Else
'                                mIntLib.Phase = 1
'                        End Select
'                    Case 2
'                        Select Case BufChar
'                            Case STX
'                                If bEndChk = True Then
'                                    RcvBuffer = ""
'                                Else
'                                    bSTXChk = True
'                                End If
'                                bEndChk = True
'                            Case vbLf
'                                If bEndChk = True Then
'                                    Call EditRcvData_One
'                                End If
'                                MSComm.Output = ACK
'                                Call mIntLib.WriteLog(ACK, ccPCLog)
'                            Case vbCr
'                                If bEndChk = True Then
'                                    Call EditRcvData_One
'                                End If
'                                MSComm.Output = ACK
'                                Call mIntLib.WriteLog(ACK, ccPCLog)
'                            Case EOT
'                                If mIntLib.State = "Q" Then
'                                    MSComm.Output = ENQ
'                                    Call mIntLib.WriteLog(ENQ, ccPCLog)
'                                    mIntLib.SndPhase = 0
'                                End If
'                                mIntLib.Phase = 3
'                            Case ENQ
'                                bSTXChk = True
'                                bEndChk = True
'                                MSComm.Output = ACK
'                                Call mIntLib.WriteLog(ACK, ccPCLog)
'                            Case NAK
'                                Call EditRcvData_One
'
'                                mIntLib.SndPhase = 0
'                                mIntLib.FrameNo = 0
'
'                                MSComm.Output = ENQ
'                                Call mIntLib.WriteLog(ENQ, ccPCLog)
'
'                            Case ETB
'                                mIntLib.IsETB = True
'                            Case Else
'                                If mIntLib.IsETB = False Then
'                                    RcvBuffer = RcvBuffer & BufChar
'                                Else
'                                    mIntLib.IsETB = False
'                                End If
'                        End Select
'                    Case 3
'                        Select Case BufChar
'                            Case ACK
'                                If mIntLib.State = "Q" Then
'                                    Call SendOrder
'                                End If
'                            Case ENQ
'                                bSTXChk = False
'                                bEndChk = True
'                                MSComm.Output = ACK
'                                Call mIntLib.WriteLog(ACK, ccPCLog)
'                                mIntLib.Phase = 2
'                            Case NAK
'                                mIntLib.SndPhase = 0
'                                mIntLib.FrameNo = 0
'                                MSComm.Output = ENQ
'                                Call mIntLib.WriteLog(ENQ, ccPCLog)
'                                mIntLib.Phase = 3
'                            Case EOT
'                                mIntLib.Phase = 1
'
'                        End Select
'                End Select
'            Next i
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
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strRackNo    As String   '수신한 Rack No
    Dim strPos       As String   '수신한 Sample Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과
    Dim strIntResult As String
    Dim strTemp      As String
    Dim strTmpBase   As String   '수신한 장비기준 검사명
    Dim strAlarm     As String
    
    Set objIntNms = mIntLib.IntNms
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        strRcvBuf = Replace(strRcvBuf, vbCr, "")
        strRcvBuf = Replace(strRcvBuf, vbLf, "")
        
        If objBuffer.Seq = "1" Then
            strType = Mid$(strRcvBuf, 2, 1)
        Else
            strType = Mid$(strRcvBuf, 1, 1)
        End If

        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "Q"    '## Request Information
                '## 바코드번호, Sample Type, Rack No, Position, Kind, Priority 조회
                strTemp = mGetP(strRcvBuf, 3, "|")
                strBarNo = Trim$(mGetP(strTemp, 3, "^"))

                With mOrder
                    .ClsClear
                    .BarNo = strBarNo
                    .SmplNo = Trim$(mGetP(strTemp, 4, "^"))
                    .RackNo = Trim$(mGetP(strTemp, 5, "^"))
                    .Pos = Trim$(mGetP(strTemp, 6, "^"))
                    .RackType = Trim$(mGetP(strTemp, 8, "^"))
                    .ContType = Trim$(mGetP(strTemp, 9, "^"))
                    .Kind = Trim$(mGetP(strTemp, 10, "^"))
                End With
                Call GetOrder(strBarNo)
                mIntLib.State = "Q"
                strQState = "Q"

            Case "O"    '## Order
                '## 바코드번호, Rack No, Sample Position 조회
                
                'O|1|11120112|0^50003^1^^S1^SC|^^^962^\^^^963^\^^^964^\^^^989^\^^^990^\^^^991^\^^^26901^\^^^28004^\^^^28005^\^^^28010^\^^^28011^\^^^28012^\^^^28013^\^^^28014^\^^^28020^\^^^28022^\^^^28024^\^^^28056^\^^^28057^\^^^28416^\^^^28421^\^^^28454^\^^^28457^|R||20200115120833||||N||||1|||||||20200115120102|||F
                'O|1|10002198328|0^50042^3^^S1^SC|^^^21002^|R||20201103082859||||N||||1|||||||20201103083156|||F
                strBarNo = mGetP(strRcvBuf, 3, "|")
                strTemp = mGetP(strRcvBuf, 4, "|")
                
                strRackNo = Trim$(mGetP(strTemp, 2, "^"))
                strPos = Trim$(mGetP(strTemp, 3, "^"))

                Set objIntInfo = New clsIISIntInfo
                With objIntInfo
                    .BarNo = strBarNo
                    .SpcPos = strPos & "/" & strRackNo
                End With
                strSpcPos = ""
                strSpcPos = strPos & "/" & strRackNo
                
            Case "R"    '## Result
                'R|22|^^^28454/|7.0|g/dL||N||F|||||C11|1|1|0|902600^^^|01909^^^
                strIntBase = mGetP(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"), 1, "/")
                strIntResult = Trim$(mGetP(strRcvBuf, 4, "|"))
                
                'RPR 일경우
                If strIntBase = "22127" Then
                    '-- 정성
                    strTmpBase = strIntBase & "C"
                    
                    strResult = strIntResult
                    If IsNumeric(strResult) Then
                        If strResult < 1 Then
                            strResult = "NonReactive"
                        Else
                            strResult = "Reactive"
                        End If
                    End If
                    
                    If objIntNms.ExistIntBase(strTmpBase) Then
                        Call objIntInfo.IntResults.Add(strTmpBase, objIntNms.GetIntNm(strTmpBase), strResult, strResult, strSpcPos)
                    End If
                    
                    '-- 정량
                    strTmpBase = strIntBase & "N"
                    strResult = strIntResult
                    
                    If objIntNms.ExistIntBase(strTmpBase) Then
                        Call objIntInfo.IntResults.Add(strTmpBase, objIntNms.GetIntNm(strTmpBase), strResult, strResult, strSpcPos)
                    End If
                    
                Else
                    strResult = strIntResult
                    
                    If objIntNms.ExistIntBase(strIntBase) Then
                        Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strResult, strResult, strSpcPos)
                    End If
                End If
                
                mIntLib.State = "R"
            Case "C"
                strAlarm = Trim$(mGetP(strRcvBuf, 4, "|"))
                strAlarm = ConvertDataAlarmCode(strAlarm)
                
                If strAlarm = "72" Then
                    Call mIntErrors.AddX("S011", mEqpCd, mEqpKey, mOrder.Pos & "/" & mOrder.RackNo, strAlarm)
                End If
                
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

Private Sub EditRcvData_One()
    Dim objIntInfo   As clsIISIntInfo    '인터페이스 검체정보 클래스
    Dim objIntNms    As clsIISIntNms     '장비별 검사항목 컬렉션 클래스
    'Dim objBuffer    As clsIISBuffer     '버퍼클래스

    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strRackNo    As String   '수신한 Rack No
    Dim strPos       As String   '수신한 Sample Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과
    Dim strIntResult As String
    Dim strTemp      As String
    Dim strTmpBase   As String   '수신한 장비기준 검사명
    Dim strAlarm     As String
    Dim ii           As Integer
    
    Set objIntNms = mIntLib.IntNms
    'For Each objBuffer In mIntLib.Buffers
        strRcvBuf = RcvBuffer
        
        ii = InStr(1, RcvBuffer, "|")
        If ii <> 0 Then
            strType = Mid(RcvBuffer, ii - 1, 1)
        Else
            Exit Sub
        End If

        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "Q"    '## Request Information
                '## 바코드번호, Sample Type, Rack No, Position, Kind, Priority 조회
                strTemp = mGetP(strRcvBuf, 3, "|")
                strBarNo = Trim$(mGetP(strTemp, 3, "^"))

                With mOrder
                    .ClsClear
                    .BarNo = strBarNo
                    .SmplNo = Trim$(mGetP(strTemp, 4, "^"))
                    .RackNo = Trim$(mGetP(strTemp, 5, "^"))
                    .Pos = Trim$(mGetP(strTemp, 6, "^"))
                    .RackType = Trim$(mGetP(strTemp, 8, "^"))
                    .ContType = Trim$(mGetP(strTemp, 9, "^"))
                    .Kind = Trim$(mGetP(strTemp, 10, "^"))
                End With
                Call GetOrder(strBarNo)
                mIntLib.State = "Q"
                strQState = "Q"

            Case "O"    '## Order
                '## 바코드번호, Rack No, Sample Position 조회
                
                'O|1|11120112|0^50003^1^^S1^SC|^^^962^\^^^963^\^^^964^\^^^989^\^^^990^\^^^991^\^^^26901^\^^^28004^\^^^28005^\^^^28010^\^^^28011^\^^^28012^\^^^28013^\^^^28014^\^^^28020^\^^^28022^\^^^28024^\^^^28056^\^^^28057^\^^^28416^\^^^28421^\^^^28454^\^^^28457^|R||20200115120833||||N||||1|||||||20200115120102|||F
                'O|1|10002198328|0^50042^3^^S1^SC|^^^21002^|R||20201103082859||||N||||1|||||||20201103083156|||F
                strBarNo = mGetP(strRcvBuf, 3, "|")
                strTemp = mGetP(strRcvBuf, 4, "|")
                
                strRackNo = Trim$(mGetP(strTemp, 2, "^"))
                strPos = Trim$(mGetP(strTemp, 3, "^"))

                Set objIntInfo = New clsIISIntInfo
                With objIntInfo
                    .BarNo = strBarNo
                    .SpcPos = strPos & "/" & strRackNo
                End With
                strSpcPos = ""
                strSpcPos = strPos & "/" & strRackNo
                
            Case "R"    '## Result
                'R|22|^^^28454/|7.0|g/dL||N||F|||||C11|1|1|0|902600^^^|01909^^^
                strIntBase = mGetP(mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^"), 1, "/")
                strIntResult = Trim$(mGetP(strRcvBuf, 4, "|"))
                
                'RPR 일경우
                If strIntBase = "28073" Then
                    '-- 정성
                    strTmpBase = strIntBase & "C"
                    
                    strResult = strIntResult
                    If strResult < 1 Then
                        strResult = "NonReactive"
                    Else
                        strResult = "Reactive"
                    End If
                    
                    If objIntNms.ExistIntBase(strTmpBase) Then
                        Call objIntInfo.IntResults.Add(strTmpBase, objIntNms.GetIntNm(strTmpBase), strResult, strResult, strSpcPos)
                    End If
                    
                    '-- 정량
                    strTmpBase = strIntBase & "N"
                    strResult = strIntResult
                    
                    If objIntNms.ExistIntBase(strTmpBase) Then
                        Call objIntInfo.IntResults.Add(strTmpBase, objIntNms.GetIntNm(strTmpBase), strResult, strResult, strSpcPos)
                    End If
                    
                Else
                    strResult = strIntResult
                    
                    If objIntNms.ExistIntBase(strIntBase) Then
                        Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), strResult, strResult, strSpcPos)
                    End If
                End If
                
                mIntLib.State = "R"
            Case "C"
                strAlarm = Trim$(mGetP(strRcvBuf, 4, "|"))
                strAlarm = ConvertDataAlarmCode(strAlarm)
                
                '72 : sample clot
                If strAlarm = "72" Then
                    Call mIntErrors.AddX("S011", mEqpCd, mEqpKey, mOrder.Pos & "/" & mOrder.RackNo, strAlarm)
                End If
                
            Case "L"    '## Terminator
                '## DB에 결과저장
                If mIntLib.State = "R" Then
                    Call SaveServer(objIntInfo)
                    Set objIntInfo = Nothing
                    mIntLib.State = ""
                End If
        End Select
    'Next
    Set objIntNms = Nothing
    'Set objBuffer = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 결과판정, 결과저장, 화면표시
'   인수 :
'       - pIntInfo : 인터페이스 검체정보 클래스
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pIntInfo As clsIISIntInfo)
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim vBarNo      As Variant 'Spread의 바코드번호
    Dim strBarNo    As String  '바코드번호
    Dim strSpcYy    As String  '검체연도
    Dim lngSpcNo    As Long    '검체번호
    Dim i           As Long

    Me.MousePointer = vbHourglass

    strBarNo = pIntInfo.BarNo
    'objAccInfo.SpcPos = pIntInfo.SpcPos
    
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

        Call SetComplete2(objAccInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Set objAccInfo = Nothing

        '## ClientDb, Server에 결과저장
        Call mIntLib.SaveClientDb(strSpcYy, lngSpcNo)
        Call mIntLib.SaveResult(strSpcYy, lngSpcNo)
        Call mIntLib.Remove(strSpcYy, lngSpcNo)
        StatusBar.Panels(2).Text = "검체번호:" & strBarNo & " 를 정상적으로 결과저장 했습니다."
    End If

    '## tblReady에서 전송된 검체삭제
    With tblReady
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
            If CStr(vBarNo) = strBarNo Then
                Call .DeleteRows(i, 1)
                Exit For
            End If
        Next i
    End With

    Me.MousePointer = vbDefault
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim strOutput  As String            '송신할 데이터

    If pBarNo = "*************" Then    '## 바코드리딩 실패시
        cmdAlarm.BackColor = vbRed
        Call mIntErrors.AddX("X002", mEqpCd, mEqpKey, mOrder.Pos & "/" & mOrder.RackNo, _
             "Sample ID Read Error.")
        mOrder.NoOrder = True
    Else
        Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
        If Not (objAccInfo Is Nothing) Then
            '## tblReady, tblresult, Label에 정보표시
            Call SetReady(objAccInfo)
            Call SetLabel(objAccInfo)
            Call SetResult(objAccInfo)
    
            mOrder.AccInfo = objAccInfo
            mOrder.NoOrder = False
            Set objAccInfo = Nothing
        Else
            mOrder.NoOrder = True
        End If
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '송신할 데이터

    Select Case mIntLib.SndPhase
        Case -1     '## 모든 오더 전송을 완료한 경우
            strOutput = EOT
            MSComm.Output = strOutput
            Call mIntLib.WriteLog(strOutput, ccPCLog)
            mIntLib.State = ""
            strQState = ""
            Exit Sub
            
        Case 0      '## 최초 오더 전송인 경우
            strOutput = ""
            strOutput = strOutput & "H|\^&|||host_COMPUTER|||||LST008AS|TSDWN^REPLY|P|1" & vbCr
            strOutput = strOutput & "P|1|||||||||||||" & vbCr
            
            If mOrder.NoOrder = False Then
                '-- 응급Rack
                If Mid(mOrder.RackNo, 1, 4) = "4000" Then
                    strOutput = strOutput & "O|1|"
                    strOutput = strOutput & mOrder.BarNo & "|"
                    strOutput = strOutput & mOrder.SmplNo & "^" & mOrder.RackNo & "^" & mOrder.Pos & "^^"
                    strOutput = strOutput & mOrder.RackType & "^" & mOrder.ContType & "|"
                    strOutput = strOutput & mOrder.GetOrder & "|"
                    strOutput = strOutput & "S||"
                    strOutput = strOutput & Format(Now, "yyyymmddhhmmss") & "||||A||||1||||||||||O" & vbCr
                Else
                    strOutput = strOutput & "O|1|"
                    strOutput = strOutput & mOrder.BarNo & "|"
                    strOutput = strOutput & mOrder.SmplNo & "^" & mOrder.RackNo & "^" & mOrder.Pos & "^^"
                    strOutput = strOutput & mOrder.RackType & "^" & mOrder.ContType & "|"
                    strOutput = strOutput & mOrder.GetOrder & "|"
                    strOutput = strOutput & "R||"
                    strOutput = strOutput & Format(Now, "yyyymmddhhmmss") & "||||A||||"
                    'S1: Serum, S2: Urine, S3: Plasma, S4: CSF, S5: Other
                    '-- Serum Rack
                    If mOrder.RackType = "S1" Then
                        strOutput = strOutput & "1"
                    '-- Urine Rack
                    ElseIf mOrder.RackType = "S2" Then
                        strOutput = strOutput & "2"
                    Else
                        strOutput = strOutput & "1"
                    End If
                    strOutput = strOutput & "||||||||||O" & vbCr
                End If
            Else
                '## 접수정보가 없는경우: 검사항목 정보를 보내지 않음!
                strOutput = strOutput & "O|1|"
                strOutput = strOutput & mOrder.BarNo & "|"
                strOutput = strOutput & mOrder.SmplNo & "^" & mOrder.RackNo & "^" & mOrder.Pos & "^^"
                strOutput = strOutput & mOrder.RackType & "^" & mOrder.ContType & "|"
                strOutput = strOutput & "|"
                strOutput = strOutput & "R||"
                '2020-11-10
                'strOutput = strOutput & Format(Now, "yyyymmddhhmmss") & "||||N||||1||||||||||O" & vbCr
                strOutput = strOutput & Format(Now, "yyyymmddhhmmss") & "||||C||||"
                '-- Serum Rack
                If mOrder.RackType = "S1" Then
                    strOutput = strOutput & "1"
                '-- Urine Rack
                ElseIf mOrder.RackType = "S2" Then
                    strOutput = strOutput & "2"
                Else
                    strOutput = strOutput & "1"
                End If
                strOutput = strOutput & "||||||||||O" & vbCr
            End If
            strOutput = strOutput & "C|1|I||G" & vbCr
            strOutput = strOutput & "L|1|N" & vbCr
            strOutput = mIntLib.GetFrameNo & strOutput
            
        Case 1      '## 오더 전송후 남은 문자열이 있는 경우
            strOutput = mIntLib.GetFrameNo & mOrder.Order
    End Select
    
    If Len(strOutput) >= 230 Then
        mOrder.Order = Mid$(strOutput, 231)
        strOutput = Mid$(strOutput, 1, 230) & ETB
        mIntLib.SndPhase = 1
    Else
        strOutput = strOutput & ETX
        mIntLib.SndPhase = -1
    End If
    
    strOutput = STX & strOutput & mOrder.GetChkSum(strOutput) & vbCrLf
    MSComm.Output = strOutput
    Call mIntLib.WriteLog(strOutput, ccPCLog)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder_wSck()
    Dim strOutput As String     '송신할 데이터

    Select Case mIntLib.SndPhase
        Case -1     '## 모든 오더 전송을 완료한 경우
            strOutput = EOT
            'MSComm.Output = strOutput
            Call wSck.SendData(strOutput)
            Call mIntLib.WriteLog(strOutput, ccPCLog)
            mIntLib.State = ""
            strQState = ""
            Exit Sub
            
        Case 0      '## 최초 오더 전송인 경우
            strOutput = ""
            strOutput = strOutput & "H|\^&|||host_COMPUTER|||||LST008AS|TSDWN^REPLY|P|1" & vbCr
            strOutput = strOutput & "P|1|||||||||||||" & vbCr
            
            If mOrder.NoOrder = False Then
                '-- 응급Rack
                If Mid(mOrder.RackNo, 1, 4) = "4000" Then
                    strOutput = strOutput & "O|1|"
                    strOutput = strOutput & mOrder.BarNo & "|"
                    strOutput = strOutput & mOrder.SmplNo & "^" & mOrder.RackNo & "^" & mOrder.Pos & "^^"
                    strOutput = strOutput & mOrder.RackType & "^" & mOrder.ContType & "|"
                    strOutput = strOutput & mOrder.GetOrder & "|"
                    strOutput = strOutput & "S||"
                    strOutput = strOutput & Format(Now, "yyyymmddhhmmss") & "||||A||||1||||||||||O" & vbCr
                Else
                    strOutput = strOutput & "O|1|"
                    strOutput = strOutput & mOrder.BarNo & "|"
                    strOutput = strOutput & mOrder.SmplNo & "^" & mOrder.RackNo & "^" & mOrder.Pos & "^^"
                    strOutput = strOutput & mOrder.RackType & "^" & mOrder.ContType & "|"
                    strOutput = strOutput & mOrder.GetOrder & "|"
                    strOutput = strOutput & "R||"
                    strOutput = strOutput & Format(Now, "yyyymmddhhmmss") & "||||A||||"
                    'S1: Serum, S2: Urine, S3: Plasma, S4: CSF, S5: Other
                    '-- Serum Rack
                    If mOrder.RackType = "S1" Then
                        strOutput = strOutput & "1"
                    '-- Urine Rack
                    ElseIf mOrder.RackType = "S2" Then
                        strOutput = strOutput & "2"
                    Else
                        strOutput = strOutput & "1"
                    End If
                    strOutput = strOutput & "||||||||||O" & vbCr
                End If
            Else
                '## 접수정보가 없는경우: 검사항목 정보를 보내지 않음!
                strOutput = strOutput & "O|1|"
                strOutput = strOutput & mOrder.BarNo & "|"
                strOutput = strOutput & mOrder.SmplNo & "^" & mOrder.RackNo & "^" & mOrder.Pos & "^^"
                strOutput = strOutput & mOrder.RackType & "^" & mOrder.ContType & "|"
                strOutput = strOutput & "|"
                strOutput = strOutput & "R||"
                strOutput = strOutput & Format(Now, "yyyymmddhhmmss") & "||||N||||1||||||||||O" & vbCr
            End If
            strOutput = strOutput & "C|1|I||G" & vbCr
            strOutput = strOutput & "L|1|N" & vbCr
            strOutput = mIntLib.GetFrameNo & strOutput
            
        Case 1      '## 오더 전송후 남은 문자열이 있는 경우
            strOutput = mIntLib.GetFrameNo & mOrder.Order
    End Select
    
    If Len(strOutput) >= 230 Then
        mOrder.Order = Mid$(strOutput, 231)
        strOutput = Mid$(strOutput, 1, 230) & ETB
        mIntLib.SndPhase = 1
    Else
        strOutput = strOutput & ETX
        mIntLib.SndPhase = -1
    End If
    
    strOutput = STX & strOutput & mOrder.GetChkSum(strOutput) & vbCrLf
    'MSComm.Output = strOutput
    Call wSck.SendData(strOutput)
    Call mIntLib.WriteLog(strOutput, ccPCLog)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblReady에 정보표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetReady(ByVal pAccInfo As clsIISAccInfo)
    With tblReady
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TReadyEnum.ccNo:     .Text = mOrder.Pos & "/" & mOrder.RackNo
        .Col = TReadyEnum.ccBarNo:  .Text = pAccInfo.GetBarNo
        .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)

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

        .Col = TCompleteEnum.ccNo:      .Text = strSpcPos 'pAccInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pAccInfo.GetBarNo
        .Col = TCompleteEnum.ccAccNo:   .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)

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
                '.Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                        DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & DIV & _
                        IIf(objResult.Ref.RefFg = "1", mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal), "") & DIV & _
                        objResult.IntInfo
                .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                        DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & "" & DIV & _
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
            Error.SetLog App.EXEName, "frmIISHitachi7600", "GetEqpComm", strErrMsg, Now
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
        Error.SetLog App.EXEName, "frmIISHitachi7600", "GetEqpComm", strErrMsg, Now
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

    If pFlag = ccAll Then
        Call mTblClear(tblReady):   Call mTblClear(tblComplete)
        Call mTblClear(tblResult)
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

Private Sub tmrQ_Timer()
    
    tmrQ.Enabled = False
    
    If strQState = "Q" Then
        mIntLib.SndPhase = 0
        mIntLib.FrameNo = 0
        MSComm.Output = ENQ
        Call mIntLib.WriteLog(ENQ, ccPCLog)
    End If
    
End Sub

Private Function ConvertDataAlarmCode(ByVal Scode As String) As String
    
    Dim sTmp    As String
    
    ConvertDataAlarmCode = "": sTmp = ""
    
    Select Case Trim(Scode)
        Case "0": sTmp = ""
        Case "1": sTmp = "ADC.E"
        Case "2": sTmp = ">Cuvet"
        Case "3": sTmp = "Samp.S"
        Case "4": sTmp = "Reag.S"
        Case "5": sTmp = ">ABS"
        Case "6": sTmp = ">Proz"
        Case "7": sTmp = ">Reac0"
        Case "8": sTmp = ">Reac1"
        Case "9": sTmp = ">Reac2"
        Case "10": sTmp = ">Lin"
        Case "11": sTmp = ">Lin"
        Case "12": sTmp = "S1A.E"
        Case "13": sTmp = "Dup.E"
        Case "14": sTmp = "Std.E"
        Case "15": sTmp = "Sens.E"
        Case "16": sTmp = "Cal.E"
        Case "17": sTmp = "SD.E"
        Case "18": sTmp = "ISE.N"
        Case "19": sTmp = "ISE.E"
        Case "20": sTmp = "Slop.E"
        Case "21": sTmp = "Prep.E"
        Case "22": sTmp = "IStd.E"
        Case "23": sTmp = "<>Test"
        Case "24": sTmp = "CmpT.E"
        Case "25": sTmp = "CmpT.?"
        Case "26": sTmp = ">Test"
        Case "27": sTmp = "<Test"
        Case "28": sTmp = "R4SD"
        Case "29": sTmp = "S2-2Sa"
        Case "30": sTmp = "S2-2Sw"
        Case "31": sTmp = "S4-1Sa"
        Case "32": sTmp = "S4-1S4"
        Case "33": sTmp = "S10Xa"
        Case "34": sTmp = "S10Xw"
        Case "35": sTmp = "Q3SD"
        Case "36": sTmp = "Q2.5SD"
        Case "37": sTmp = "ClcT.E"
        Case "38": sTmp = "Over.E"
        Case "39": sTmp = "Calc.?"
        Case "40": sTmp = "H"
        Case "41": sTmp = "L"
        Case "42": sTmp = "Edited"
        Case "43": sTmp = "Cal.E"
        Case "44": sTmp = ">Rept"
        Case "45": sTmp = "<rept"
        Case "46": sTmp = ""
        Case "47": sTmp = ""
        Case "48": sTmp = "QCH"
        Case "49": sTmp = "QCL"
        Case "50": sTmp = ""
        Case "51": sTmp = "Rsp1.E"
        Case "52": sTmp = "Rsp2.E"
        Case "53": sTmp = "Cond.E"
        Case "54": sTmp = "S2Pr.E"
        Case "55": sTmp = ""
        Case "56": sTmp = ">Kin"
        Case "57": sTmp = ""
        Case "58": sTmp = ""
        Case "59": sTmp = "MIXSTP"
        Case "60": sTmp = "MIXLOW"
        Case "61": sTmp = "Samp.V"
        Case "62": sTmp = ""
        Case "63": sTmp = ""
        Case "64": sTmp = ""
        Case "65": sTmp = ""
        Case "66": sTmp = ""
        Case "67": sTmp = ""
        Case "68": sTmp = ""
        Case "69": sTmp = ""
        Case "70": sTmp = ""
        Case "71": sTmp = ""
        Case "72": sTmp = "Smp.C"
        Case "73": sTmp = "Det.C"
        Case "74": sTmp = ""
        Case "75": sTmp = ""
        Case "76": sTmp = ""
        Case "77": sTmp = ""
        Case "78": sTmp = ""
        Case "79": sTmp = ""
        Case "80": sTmp = ""
        Case "81": sTmp = ""
        Case "82": sTmp = ""
        Case "83": sTmp = "Samp.O"
        '84 to 100
        Case "101": sTmp = "ReagEx"
        Case "102": sTmp = ""
        Case "103": sTmp = ">I.L"
        Case "104": sTmp = ">I.H"
        Case "105": sTmp = ">I.I"
        Case "106": sTmp = ">I.LH"
        Case "107": sTmp = ">I.LI"
        Case "108": sTmp = ">I.HI"
        Case "109": sTmp = ">I.LHI"
        Case "110": sTmp = ""
        Case "111": sTmp = ""
        Case "112": sTmp = ""
        Case "113": sTmp = ">A.Dif"
        Case Else: sTmp = ""
    End Select
    
    ConvertDataAlarmCode = Trim(sTmp)
    
End Function

Private Sub wSck_Close()
    
    If wSck.State <> sckClosed Then
        wSck.Close
    End If
    wSck.LocalPort = CInt(gSckPort)
    wSck.Listen

    StatusBar.Panels(2).Text = ""

End Sub

Private Sub wSck_ConnectionRequest(ByVal requestID As Long)
    
    If wSck.State <> sckClosed Then
        wSck.Close
    End If
    
    wSck.Accept requestID
    
    StatusBar.Panels(2).Text = mEqpKey & " : " & gSckPort & "번 포트에 연결되었습니다"

End Sub

Private Sub wSck_DataArrival(ByVal bytesTotal As Long)
    Dim strRcvBuffer As String
   
    wSck.GetData strRcvBuffer

    Call RcvSocketData(strRcvBuffer)

End Sub

Private Sub RcvSocketData(ByVal pRcvData As String)
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    Buffer = pRcvData
    
    Call mIntLib.WriteLog(Buffer, ccEqp)

    'Debug.Print Buffer
    
    lngBufLen = Len(Buffer)
    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)

        Select Case mIntLib.Phase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Call mIntLib.ClearBuffer
                        mIntLib.Phase = 2
                        Call wSck.SendData(ACK)
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                    Case ACK
                        If mIntLib.State = "Q" Then
                            Call SendOrder_wSck
                        Else    '-- Edit by Sewon,Oh(2006.08.02)
                            Call wSck.SendData(ACK)
                        End If
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Call mIntLib.ClearBuffer
                        wSck.SendData (ACK)
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                    Case STX
                        mIntLib.BufCnt = 1
                        Call mIntLib.ClearBuffer
                    Case ETB
                        mIntLib.IsETB = True
                        mIntLib.Phase = 3
                    Case ETX
                        mIntLib.BufCnt = mIntLib.BufCnt + 1
                        mIntLib.Phase = 3
                    Case EOT
                        mIntLib.Phase = 1
                    Case vbCr, vbLf
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
                        mIntLib.Phase = 4
                        Call wSck.SendData(ACK)
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        mIntLib.Phase = 2
                    Case EOT
                        Call EditRcvData
                        If mIntLib.State = "Q" Then
                            mIntLib.SndPhase = 1
                            mIntLib.FrameNo = 0
                            Call wSck.SendData(ENQ)
                            Call mIntLib.WriteLog(ENQ, ccPCLog)
                        End If
                        mIntLib.Phase = 1
                End Select
        End Select
    Next i
End Sub



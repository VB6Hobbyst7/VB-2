VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIISGeneXpert 
   BackColor       =   &H00DBE6E6&
   Caption         =   "GeneXpert"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   4260
      TabIndex        =   27
      Top             =   180
      Visible         =   0   'False
      Width           =   1155
   End
   Begin IISGeneXpert.sckStringData sck 
      Height          =   300
      Left            =   8940
      TabIndex        =   26
      Top             =   8730
      Visible         =   0   'False
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   529
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8070
      Top             =   8550
      _ExtentX        =   741
      _ExtentY        =   741
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
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      Height          =   495
      Left            =   11475
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   12694
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "종 료(&X)"
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
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISGeneXpert.frx":0000
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4410
      Left            =   105
      TabIndex        =   22
      Top             =   4663
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
      SpreadDesigner  =   "frmIISGeneXpert.frx":0477
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
      SpreadDesigner  =   "frmIISGeneXpert.frx":0C26
      TextTip         =   2
   End
End
Attribute VB_Name = "frmIISGeneXpert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISGeneXpert.frm
'   작성자  : 오세원
'   내  용  : GeneXpert 장비폼
'   작성일  : 2015-05-28
'   버  전  :
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

'## Datalog Length 상수
Private Const RACKNOLEN     As Long = 4
Private Const CUPPOSLEN     As Long = 2
Private Const SAMPLENOLEN   As Long = 4
Private Const TESTCDLEN     As Long = 2
Private Const RESULTLEN     As Long = 6
Private Const FLAGSLEN      As Long = 2

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

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property

Private Sub Command1_Click()
    Dim sBuf As String
    
           sBuf = "1H|@^\|URM-T2HXZ2UA-01||GeneXpert PC^GeneXpert^4.4a|||||1||P|1394-97|20150528161125" & vbCr
    sBuf = sBuf & "P|1|||160019221978" & vbCr
    sBuf = sBuf & "O|1|160019221978||^^^CAR|R|20161219080405|||||||||ORH||||||||||F" & vbCr
    sBuf = sBuf & "R|1|^CAR^^11^Xpert Carba-R^2^IMP1^|IMP1 NOT DETECTED^|||||F||<None>|2016171" & vbCr

'R|1|^CAR^^11^Xpert Carba-R^2^IMP1^|IMP1 NOT DETECTED^|||||F||<None>|2016171
'2219080405|20161219085358|Cepheid-412BW35^802894^621327^476582344^06204^20171022
'R|2|^CAR^^11^^^IMP1^|NEG^|||
'R|3|^CAR^^11^^^IMP1^Ct|^0|||
'R|4|^CAR^^11^^^IMP1^EndPt|^1.0|||
'R|5|^CAR^^11^^^SPC^|PASS^|||
'R|6|^CAR^^11^^^SPC^Ct|^35.6|||
'R|7|^CAREF


    sBuf = sBuf & "2219080405|20161219085358|Cepheid-412BW35^802894^621327^476582344^06204^20171022" & vbCr
    sBuf = sBuf & "R|2|^RIF^^1^^^Probe D^|NEG^|||" & vbCr
    sBuf = sBuf & "R|3|^RIF^^1^^^Probe D^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|4|^RIF^^1^^^Probe D^EndPt|^0|||" & vbCr
    sBuf = sBuf & "R|5|^RIF^^1^^^Probe C^|NEG^|||" & vbCr
    sBuf = sBuf & "R|6|^RIF^^1^^^Probe C^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|7|^RIF5B" & vbCr

    sBuf = sBuf & "3^^1^^^Probe C^EndPt|^2.0|||" & vbCr
    sBuf = sBuf & "R|8|^RIF^^1^^^Probe E^|NEG^|||" & vbCr
    sBuf = sBuf & "R|9|^RIF^^1^^^Probe E^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|10|^RIF^^1^^^Probe E^EndPt|^-2.0|||" & vbCr
    sBuf = sBuf & "R|11|^RIF^^1^^^Probe B^|NEG^|||" & vbCr
    sBuf = sBuf & "R|12|^RIF^^1^^^Probe B^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|13|^RIF^^1^^^Probe B^EndPt|^5.0|||" & vbCr
    sBuf = sBuf & "R|14|^RIF^^D8" & vbCr

    sBuf = sBuf & "41^^^Probe A^|NEG^|||" & vbCr
    sBuf = sBuf & "R|15|^RIF^^1^^^Probe A^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|16|^RIF^^1^^^Probe A^EndPt|^-1.0|||" & vbCr
    sBuf = sBuf & "R|17|^RIF^^1^^^SPC^|PASS^|||" & vbCr
    sBuf = sBuf & "R|18|^RIF^^1^^^SPC^Ct|^25.5|||" & vbCr
    sBuf = sBuf & "R|19|^RIF^^1^^^SPC^EndPt|^229.0|||" & vbCr
    sBuf = sBuf & "R|20|^RIF^^3^Xpert MTB-RIF Assay G4^5^Rif Resistance^|9C" & vbCr

    sBuf = sBuf & "5^|||||F||<None>|20150518102153|20150518120349|Cepheid-412BW35^802426^626760^135318629^07810^20161023" & vbCr
    sBuf = sBuf & "R|21|^RIF^^3^^^Probe D^|NEG^|||" & vbCr
    sBuf = sBuf & "R|22|^RIF^^3^^^Probe D^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|23|^RIF^^3^^^Probe D^EndPt|^0|||" & vbCr
    sBuf = sBuf & "R|24|^RIF^^3^^^Probe C^|NEG^|||" & vbCr
    sBuf = sBuf & "R|25|^RI0A" & vbCr
    
    sBuf = sBuf & "6F^^3^^^Probe C^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|26|^RIF^^3^^^Probe C^EndPt|^2.0|||" & vbCr
    sBuf = sBuf & "R|27|^RIF^^3^^^Probe E^|NEG^|||" & vbCr
    sBuf = sBuf & "R|28|^RIF^^3^^^Probe E^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|29|^RIF^^3^^^Probe E^EndPt|^-2.0|||" & vbCr
    sBuf = sBuf & "R|30|^RIF^^3^^^Probe B^|NEG^|||" & vbCr
    sBuf = sBuf & "R|31|^RIF^^3^^^Probe B^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|32|^RIF^^3^AC" & vbCr
    
    sBuf = sBuf & "7^^Probe B^EndPt|^5.0|||" & vbCr
    sBuf = sBuf & "R|33|^RIF^^3^^^Probe A^|NEG^|||" & vbCr
    sBuf = sBuf & "R|34|^RIF^^3^^^Probe A^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|35|^RIF^^3^^^Probe A^EndPt|^-1.0|||" & vbCr
    sBuf = sBuf & "R|36|^RIF^^3^^^SPC^|PASS^|||" & vbCr
    sBuf = sBuf & "R|37|^RIF^^3^^^SPC^Ct|^25.5|||" & vbCr
    sBuf = sBuf & "R|38|^RIF^^3^^^SPC^EndPt|^229.0|||" & vbCr
    sBuf = sBuf & "R|39|^RIF^^2^Xpert D7" & vbCr
    
    sBuf = sBuf & "0MTB-RIF Assay G4^5^QC Check^|^|||||F||<None>|20150518102153|20150518120349|Cepheid-412BW35^802426^626760^135318629^07810^20161023" & vbCr
    sBuf = sBuf & "R|40|^RIF^^2^^^QC-1^|NEG^|||" & vbCr
    sBuf = sBuf & "R|41|^RIF^^2^^^QC-1^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|42|^RIF^^2^^^QC-1^EndPt|^0|||" & vbCr
    sBuf = sBuf & "L|1|N" & vbCr
    sBuf = sBuf & "7D" & vbCr
    sBuf = sBuf & ""
    
        
        
    sBuf = ""
    sBuf = sBuf & ""
    sBuf = sBuf & "1H|@^\|URM-sBgYZ2UA-02||GeneXpert PC^GeneXpert^4.4a|||||1||P|1394-97|20150528161726" & vbCr
    sBuf = sBuf & "P|1|||51" & vbCr
    sBuf = sBuf & "O|1|Xpert M 041015114144||^^^RIF|R|20150410114245|||||||||ORH||||||||||F" & vbCr
    sBuf = sBuf & "R|1|^RIF^^1^Xpert MTB-RIF Assay G4^5^MTB^|MTB DETECTED HIGH^|||||F||<None>|D8" & vbCr

    sBuf = sBuf & "220150410114245|20150410132445|Cepheid-412BW35^802426^626760^432773969^08101^20160911" & vbCr
    sBuf = sBuf & "R|2|^RIF^^1^^^Probe D^|POS^|||" & vbCr
    sBuf = sBuf & "R|3|^RIF^^1^^^Probe D^Ct|^14.4|||" & vbCr
    sBuf = sBuf & "R|4|^RIF^^1^^^Probe D^EndPt|^212.0|||" & vbCr
    sBuf = sBuf & "R|5|^RIF^^1^^^Probe C^|POS^|||" & vbCr
    sBuf = sBuf & "R|6|^RIF^^1^^^Probe CAC" & vbCr

    sBuf = sBuf & "3^Ct|^14.0|||" & vbCr
    sBuf = sBuf & "R|7|^RIF^^1^^^Probe C^EndPt|^175.0|||" & vbCr
    sBuf = sBuf & "R|8|^RIF^^1^^^Probe E^|POS^|||" & vbCr
    sBuf = sBuf & "R|9|^RIF^^1^^^Probe E^Ct|^15.0|||" & vbCr
    sBuf = sBuf & "R|10|^RIF^^1^^^Probe E^EndPt|^102.0|||" & vbCr
    sBuf = sBuf & "R|11|^RIF^^1^^^Probe B^|POS^|||" & vbCr
    sBuf = sBuf & "R|12|^RIF^^1^^^Probe B^Ct|^14.8|||" & vbCr
    sBuf = sBuf & "R|13|^RIF^^1^^^ProCF" & vbCr

    sBuf = sBuf & "4be B^EndPt|^105.0|||" & vbCr
    sBuf = sBuf & "R|14|^RIF^^1^^^Probe A^|POS^|||" & vbCr
    sBuf = sBuf & "R|15|^RIF^^1^^^Probe A^Ct|^13.5|||" & vbCr
    sBuf = sBuf & "R|16|^RIF^^1^^^Probe A^EndPt|^104.0|||" & vbCr
    sBuf = sBuf & "R|17|^RIF^^1^^^SPC^|NA^|||" & vbCr
    sBuf = sBuf & "R|18|^RIF^^1^^^SPC^Ct|^27.4|||" & vbCr
    sBuf = sBuf & "R|19|^RIF^^1^^^SPC^EndPt|^230.0|||" & vbCr
    sBuf = sBuf & "R|20|^RIF^^3^Xpert MB1" & vbCr

    sBuf = sBuf & "5TB-RIF Assay G4^5^Rif Resistance^|Rif Resistance NOT DETECTED^|||||F||<None>|20150410114245|20150410132445|Cepheid-412BW35^802426^626760^432773969^08101^20160911" & vbCr
    sBuf = sBuf & "R|21|^RIF^^3^^^Probe D^|POS^|||" & vbCr
    sBuf = sBuf & "R|22|^RIF^^3^^^Probe D^Ct|^14.4|||" & vbCr
    sBuf = sBuf & "R|23|^RIF^^FC" & vbCr

    sBuf = sBuf & "63^^^Probe D^EndPt|^212.0|||" & vbCr
    sBuf = sBuf & "R|24|^RIF^^3^^^Probe C^|POS^|||" & vbCr
    sBuf = sBuf & "R|25|^RIF^^3^^^Probe C^Ct|^14.0|||" & vbCr
    sBuf = sBuf & "R|26|^RIF^^3^^^Probe C^EndPt|^175.0|||" & vbCr
    sBuf = sBuf & "R|27|^RIF^^3^^^Probe E^|POS^|||" & vbCr
    sBuf = sBuf & "R|28|^RIF^^3^^^Probe E^Ct|^15.0|||" & vbCr
    sBuf = sBuf & "R|29|^RIF^^3^^^Probe E^EndPt|^102.0|||" & vbCr
    sBuf = sBuf & "55" & vbCr

    sBuf = sBuf & "7R|30|^RIF^^3^^^Probe B^|POS^|||" & vbCr
    sBuf = sBuf & "R|31|^RIF^^3^^^Probe B^Ct|^14.8|||" & vbCr
    sBuf = sBuf & "R|32|^RIF^^3^^^Probe B^EndPt|^105.0|||" & vbCr
    sBuf = sBuf & "R|33|^RIF^^3^^^Probe A^|POS^|||" & vbCr
    sBuf = sBuf & "R|34|^RIF^^3^^^Probe A^Ct|^13.5|||" & vbCr
    sBuf = sBuf & "R|35|^RIF^^3^^^Probe A^EndPt|^104.0|||" & vbCr
    sBuf = sBuf & "R|36|^RIF^^3^^^SPC^|NA^|||" & vbCr
    sBuf = sBuf & "R86" & vbCr

    sBuf = sBuf & "0|37|^RIF^^3^^^SPC^Ct|^27.4|||" & vbCr
    sBuf = sBuf & "R|38|^RIF^^3^^^SPC^EndPt|^230.0|||" & vbCr
    sBuf = sBuf & "R|39|^RIF^^2^Xpert MTB-RIF Assay G4^5^QC Check^|^|||||F||<None>|20150410114245|20150410132445|Cepheid-412BW35^802426^626760^432773969^08101^20160911" & vbCr
    sBuf = sBuf & "R|40|^RIF^^2^^^QC-1^|NEG^|AE" & vbCr

    sBuf = sBuf & "1||" & vbCr
    sBuf = sBuf & "R|41|^RIF^^2^^^QC-1^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|42|^RIF^^2^^^QC-1^EndPt|^0|||" & vbCr
    sBuf = sBuf & "R|43|^RIF^^2^^^QC-2^|NEG^|||" & vbCr
    sBuf = sBuf & "R|44|^RIF^^2^^^QC-2^Ct|^0|||" & vbCr
    sBuf = sBuf & "R|45|^RIF^^2^^^QC-2^EndPt|^0|||" & vbCr
    sBuf = sBuf & "L|1|N" & vbCr
    sBuf = sBuf & "F9" & vbCr

    sBuf = sBuf & ""

'[P C] 
'1H|@^\|URM-kTgYZ2UA-03||GeneXpert PC^GeneXpert^4.4a|||||1||P|1394-97|20150528161727
'P|1|||215
'O|1|Xpert M 040815145117||^^^RIF|R|20150408145215|||||||||ORH||||||||||F
'R|1|^RIF^^1^Xpert MTB-RIF Assay G4^5^MTB^|MTB DETECTED MEDIUM^|||||F||<NonAB
'[P C] 
'2e>|20150408145215|20150408163343|Cepheid-412BW35^802426^627226^135318719^07810^20161023
'R|2|^RIF^^1^^^Probe D^|POS^|||
'R|3|^RIF^^1^^^Probe D^Ct|^20.8|||
'R|4|^RIF^^1^^^Probe D^EndPt|^234.0|||
'R|5|^RIF^^1^^^Probe C^|POS^|||
'R|6|^RIF^^1^^^Prob0B
'[P C] 
'3e C^Ct|^20.3|||
'R|7|^RIF^^1^^^Probe C^EndPt|^187.0|||
'R|8|^RIF^^1^^^Probe E^|POS^|||
'R|9|^RIF^^1^^^Probe E^Ct|^21.4|||
'R|10|^RIF^^1^^^Probe E^EndPt|^111.0|||
'R|11|^RIF^^1^^^Probe B^|POS^|||
'R|12|^RIF^^1^^^Probe B^Ct|^22.2|||
'R|13|^RIF^^1^^^63
'[P C] 
'4Probe B^EndPt|^88.0|||
'R|14|^RIF^^1^^^Probe A^|NEG^|||
'R|15|^RIF^^1^^^Probe A^Ct|^0|||
'R|16|^RIF^^1^^^Probe A^EndPt|^6.0|||
'R|17|^RIF^^1^^^SPC^|NA^|||
'R|18|^RIF^^1^^^SPC^Ct|^26.2|||
'R|19|^RIF^^1^^^SPC^EndPt|^283.0|||
'R|20|^RIF^^3^Xpert MTB-76
'[P C] 
'5RIF Assay G4^5^Rif Resistance^|Rif Resistance DETECTED^|||||F||<None>|20150408145215|20150408163343|Cepheid-412BW35^802426^627226^135318719^07810^20161023
'R|21|^RIF^^3^^^Probe D^|POS^|||
'R|22|^RIF^^3^^^Probe D^Ct|^20.8|||
'R|23|^RIF^^3^^^ProAA
'[P C] 
'6be D^EndPt|^234.0|||
'R|24|^RIF^^3^^^Probe C^|POS^|||
'R|25|^RIF^^3^^^Probe C^Ct|^20.3|||
'R|26|^RIF^^3^^^Probe C^EndPt|^187.0|||
'R|27|^RIF^^3^^^Probe E^|POS^|||
'R|28|^RIF^^3^^^Probe E^Ct|^21.4|||
'R|29|^RIF^^3^^^Probe E^EndPt|^111.0|||
'R|30|^R3C
'[P C] 
'7IF^^3^^^Probe B^|POS^|||
'R|31|^RIF^^3^^^Probe B^Ct|^22.2|||
'R|32|^RIF^^3^^^Probe B^EndPt|^88.0|||
'R|33|^RIF^^3^^^Probe A^|NEG^|||
'R|34|^RIF^^3^^^Probe A^Ct|^0|||
'R|35|^RIF^^3^^^Probe A^EndPt|^6.0|||
'R|36|^RIF^^3^^^SPC^|NA^|||
'R|37|^RIF^^3^^3A
'[P C] 
'0^SPC^Ct|^26.2|||
'R|38|^RIF^^3^^^SPC^EndPt|^283.0|||
'R|39|^RIF^^2^Xpert MTB-RIF Assay G4^5^QC Check^|^|||||F||<None>|20150408145215|20150408163343|Cepheid-412BW35^802426^627226^135318719^07810^20161023
'R|40|^RIF^^2^^^QC-1^|NEG^|||
'R|41|^RIF^BB
'[P C] 
'1^2^^^QC-1^Ct|^0|||
'R|42|^RIF^^2^^^QC-1^EndPt|^0|||
'R|43|^RIF^^2^^^QC-2^|NEG^|||
'R|44|^RIF^^2^^^QC-2^Ct|^0|||
'R|45|^RIF^^2^^^QC-2^EndPt|^0|||
'L|1|N
'A8
'[P C] 
'[P C] 
'1H|@^\|URM-1mgYZ2UA-04||GeneXpert PC^GeneXpert^4.4a|||||1||P|1394-97|20150528161728
'P|1|||MTB1
'O|1|Xpert M 040815130515||^^^RIF|R|20150408130630|||||||||ORH||||||||||F
'R|1|^RIF^^1^Xpert MTB-RIF Assay G4^5^MTB^|MTB DETECTED LOW^|||||F||<None>D3
'[P C] 
'2|20150408130630|20150408144800|Cepheid-412BW35^802426^627226^135318467^07810^20161023
'R|2|^RIF^^1^^^Probe D^|POS^|||
'R|3|^RIF^^1^^^Probe D^Ct|^23.3|||
'R|4|^RIF^^1^^^Probe D^EndPt|^180.0|||
'R|5|^RIF^^1^^^Probe C^|POS^|||
'R|6|^RIF^^1^^^Probe E3
'[P C] 
'3C^Ct|^23.3|||
'R|7|^RIF^^1^^^Probe C^EndPt|^155.0|||
'R|8|^RIF^^1^^^Probe E^|POS^|||
'R|9|^RIF^^1^^^Probe E^Ct|^24.3|||
'R|10|^RIF^^1^^^Probe E^EndPt|^99.0|||
'R|11|^RIF^^1^^^Probe B^|POS^|||
'R|12|^RIF^^1^^^Probe B^Ct|^23.3|||
'R|13|^RIF^^1^^^ProF0
'[P C] 
'4be B^EndPt|^124.0|||
'R|14|^RIF^^1^^^Probe A^|POS^|||
'R|15|^RIF^^1^^^Probe A^Ct|^22.6|||
'R|16|^RIF^^1^^^Probe A^EndPt|^100.0|||
'R|17|^RIF^^1^^^SPC^|NA^|||
'R|18|^RIF^^1^^^SPC^Ct|^30.8|||
'R|19|^RIF^^1^^^SPC^EndPt|^194.0|||
'R|20|^RIF^^3^Xpert MB6
'[P C] 
'5TB-RIF Assay G4^5^Rif Resistance^|Rif Resistance NOT DETECTED^|||||F||<None>|20150408130630|20150408144800|Cepheid-412BW35^802426^627226^135318467^07810^20161023
'R|21|^RIF^^3^^^Probe D^|POS^|||
'R|22|^RIF^^3^^^Probe D^Ct|^23.3|||
'R|23|^RIF^^F6
'[P C] 
'63^^^Probe D^EndPt|^180.0|||
'R|24|^RIF^^3^^^Probe C^|POS^|||
'R|25|^RIF^^3^^^Probe C^Ct|^23.3|||
'R|26|^RIF^^3^^^Probe C^EndPt|^155.0|||
'R|27|^RIF^^3^^^Probe E^|POS^|||
'R|28|^RIF^^3^^^Probe E^Ct|^24.3|||
'R|29|^RIF^^3^^^Probe E^EndPt|^99.0|||
'R8E
'[P C] 
'7|30|^RIF^^3^^^Probe B^|POS^|||
'R|31|^RIF^^3^^^Probe B^Ct|^23.3|||
'R|32|^RIF^^3^^^Probe B^EndPt|^124.0|||
'R|33|^RIF^^3^^^Probe A^|POS^|||
'R|34|^RIF^^3^^^Probe A^Ct|^22.6|||
'R|35|^RIF^^3^^^Probe A^EndPt|^100.0|||
'R|36|^RIF^^3^^^SPC^|NA^|||
'R|A9
'[P C] 
'037|^RIF^^3^^^SPC^Ct|^30.8|||
'R|38|^RIF^^3^^^SPC^EndPt|^194.0|||
'R|39|^RIF^^2^Xpert MTB-RIF Assay G4^5^QC Check^|^|||||F||<None>|20150408130630|20150408144800|Cepheid-412BW35^802426^627226^135318467^07810^20161023
'R|40|^RIF^^2^^^QC-1^|NEG^||B0
'[P C] 
'1|
'R|41|^RIF^^2^^^QC-1^Ct|^0|||
'R|42|^RIF^^2^^^QC-1^EndPt|^0|||
'R|43|^RIF^^2^^^QC-2^|NEG^|||
'R|44|^RIF^^2^^^QC-2^Ct|^0|||
'R|45|^RIF^^2^^^QC-2^EndPt|^0|||
'L|1|N
'7D
'[P C] 
'
    
                            
sBuf = sBuf & "1H|@^\|URM-j/g2VkVA-45||GeneXpert PC^GeneXpert^4.4a|||||1||P|1394-97|20161219155209" & vbCr
sBuf = sBuf & "P|1|||160019266177" & vbCr
sBuf = sBuf & "O|1|160019266177||^^^CAR|R|20161219135850|||||||||ORH||||||||||F" & vbCr
sBuf = sBuf & "R|1|^CAR^^11^Xpert Carba-R^2^IMP1^|INVALID^|||||F||<None>|20161219135850|1B" & vbCrLf
sBuf = sBuf & "220161219144912|Cepheid-412BW35^802894^678322^476582358^06204^20171022" & vbCr
sBuf = sBuf & "R|2|^CAR^^11^^^IMP1^|INVALID^|||" & vbCr
sBuf = sBuf & "R|3|^CAR^^11^^^IMP1^Ct|^0|||" & vbCr
sBuf = sBuf & "R|4|^CAR^^11^^^IMP1^EndPt|^2.0|||" & vbCr
sBuf = sBuf & "R|5|^CAR^^11^^^SPC^|FAIL^|||" & vbCr
sBuf = sBuf & "R|6|^CAR^^11^^^SPC^Ct|^0|||" & vbCr
sBuf = sBuf & "R|7|^CAR^^11^^^SPFC" & vbCrLf
sBuf = sBuf & "3C^EndPt|^10.0|||" & vbCr
sBuf = sBuf & "R|8|^CAR^^15^Xpert Carba-R^2^VIM^|INVALID^|||||F||<None>|20161219135850|20161219144912|Cepheid-412BW35^802894^678322^476582358^06204^20171022" & vbCr
sBuf = sBuf & "R|9|^CAR^^15^^^VIM^|INVALID^|||" & vbCr
sBuf = sBuf & "R|10|^CAR^^15^^^VIM^Ct|^0|||" & vbCr
sBuf = sBuf & "R|11|^CAR^^15^^^VIM^65" & vbCrLf
sBuf = sBuf & "4EndPt|^5.0|||" & vbCr
sBuf = sBuf & "R|12|^CAR^^15^^^SPC^|FAIL^|||" & vbCr
sBuf = sBuf & "R|13|^CAR^^15^^^SPC^Ct|^0|||" & vbCr
sBuf = sBuf & "R|14|^CAR^^15^^^SPC^EndPt|^10.0|||" & vbCr
sBuf = sBuf & "R|15|^CAR^^13^Xpert Carba-R^2^NDM^|INVALID^|||||F||<None>|20161219135850|20161219144912|Cepheid-412BW35^802894^678322^476582358^0620B2" & vbCrLf
sBuf = sBuf & "54^20171022" & vbCr
sBuf = sBuf & "R|16|^CAR^^13^^^NDM^|INVALID^|||" & vbCr
sBuf = sBuf & "R|17|^CAR^^13^^^NDM^Ct|^0|||" & vbCr
sBuf = sBuf & "R|18|^CAR^^13^^^NDM^EndPt|^1.0|||" & vbCr
sBuf = sBuf & "R|19|^CAR^^13^^^SPC^|FAIL^|||" & vbCr
sBuf = sBuf & "R|20|^CAR^^13^^^SPC^Ct|^0|||" & vbCr
sBuf = sBuf & "R|21|^CAR^^13^^^SPC^EndPt|^10.0|||" & vbCr
sBuf = sBuf & "R|22|^CAR^^12^Xpert Carba-R^2^KPC^|INVA89" & vbCrLf
sBuf = sBuf & "6LID^|||||F||<None>|20161219135850|20161219144912|Cepheid-412BW35^802894^678322^476582358^06204^20171022" & vbCr
sBuf = sBuf & "R|23|^CAR^^12^^^KPC^|INVALID^|||" & vbCr
sBuf = sBuf & "R|24|^CAR^^12^^^KPC^Ct|^0|||" & vbCr
sBuf = sBuf & "R|25|^CAR^^12^^^KPC^EndPt|^2.0|||" & vbCr
sBuf = sBuf & "R|26|^CAR^^12^^^SPC^|FAIL^|||" & vbCr
sBuf = sBuf & "R|27|^CAR^17" & vbCrLf
sBuf = sBuf & "7^12^^^SPC^Ct|^0|||" & vbCr
sBuf = sBuf & "R|28|^CAR^^12^^^SPC^EndPt|^10.0|||" & vbCr
sBuf = sBuf & "R|29|^CAR^^14^Xpert Carba-R^2^OXA48^|INVALID^|||||F||<None>|20161219135850|20161219144912|Cepheid-412BW35^802894^678322^476582358^06204^20171022" & vbCr
sBuf = sBuf & "R|30|^CAR^^14^^^OXA48^|INVALID^|||" & vbCr
sBuf = sBuf & "R|31|^CAR^^14^^^OXA48^Ct|^0|||" & vbCr
sBuf = sBuf & "R|32|^CAR^^14^^^OXA48^EndPt|^6.0|||" & vbCr
sBuf = sBuf & "R|33|^CAR^^14^^^SPC^|FAIL^|||" & vbCr
sBuf = sBuf & "R|34|^CAR^^14^^^SPC^Ct|^0|||" & vbCr
sBuf = sBuf & "R|35|^CAR^^14^^^SPC^EndPt|^10.0|||" & vbCr
sBuf = sBuf & "L|1|N" & vbCr
sBuf = sBuf & "7C" & vbCrLf
sBuf = sBuf & ""
    
    
'1H|@^\|URM-TbgmAtVA-01||GeneXpert PC^GeneXpert^4.4a|||||1||P|1394-97|20170406091127
'P|1|||170010676615
'O|1|170010676615||^^^CAR|R|20170406075547|||||||||ORH||||||||||F
'R|1|^CAR^^11^Xpert Carba-R^2^IMP1^|IMP1 NOT DETECTED^|||||F||<None>|2017088
'2406075547|20170406084542|Cepheid-412BW35^802894^612576^482935925^06305^20180114
'R|2|^CAR^^11^^^IMP1^|NEG^|||
'R|3|^CAR^^11^^^IMP1^Ct|^0|||
'R|4|^CAR^^11^^^IMP1^EndPt|^0|||
'R|5|^CAR^^11^^^SPC^|PASS^|||
'R|6|^CAR^^11^^^SPC^Ct|^32.9|||
'R|7|^CAR^^5B
'311^^^SPC^EndPt|^155.0|||
'R|8|^CAR^^15^Xpert Carba-R^2^VIM^|VIM NOT DETECTED^|||||F||<None>|20170406075547|20170406084542|Cepheid-412BW35^802894^612576^482935925^06305^20180114
'R|9|^CAR^^15^^^VIM^|NEG^|||
'R|10|^CAR^^15^^^VIM^Ct|^0|||
'R|11|^CD6
'4AR^^15^^^VIM^EndPt|^0|||
'R|12|^CAR^^15^^^SPC^|PASS^|||
'R|13|^CAR^^15^^^SPC^Ct|^32.9|||
'R|14|^CAR^^15^^^SPC^EndPt|^155.0|||
'R|15|^CAR^^13^Xpert Carba-R^2^NDM^|NDM NOT DETECTED^|||||F||<None>|20170406075547|20170406084542|Cepheid-412BW35^80283D
'594^612576^482935925^06305^20180114
'R|16|^CAR^^13^^^NDM^|NEG^|||
'R|17|^CAR^^13^^^NDM^Ct|^0|||
'R|18|^CAR^^13^^^NDM^EndPt|^-1.0|||
'R|19|^CAR^^13^^^SPC^|PASS^|||
'R|20|^CAR^^13^^^SPC^Ct|^32.9|||
'R|21|^CAR^^13^^^SPC^EndPt|^155.0|||
'R|22|^CAR^^12^88
'6Xpert Carba-R^2^KPC^|KPC NOT DETECTED^|||||F||<None>|20170406075547|20170406084542|Cepheid-412BW35^802894^612576^482935925^06305^20180114
'R|23|^CAR^^12^^^KPC^|NEG^|||
'R|24|^CAR^^12^^^KPC^Ct|^0|||
'R|25|^CAR^^12^^^KPC^EndPt|^0|||
'R|26|^CAR^^1ED
'72^^^SPC^|PASS^|||
'R|27|^CAR^^12^^^SPC^Ct|^32.9|||
'R|28|^CAR^^12^^^SPC^EndPt|^155.0|||
'R|29|^CAR^^14^Xpert Carba-R^2^OXA48^|OXA48 NOT DETECTED^|||||F||<None>|20170406075547|20170406084542|Cepheid-412BW35^802894^612576^482935925^06305^20180113E
'04
'R|30|^CAR^^14^^^OXA48^|NEG^|||
'R|31|^CAR^^14^^^OXA48^Ct|^0|||
'R|32|^CAR^^14^^^OXA48^EndPt|^2.0|||
'R|33|^CAR^^14^^^SPC^|PASS^|||
'R|34|^CAR^^14^^^SPC^Ct|^32.9|||
'R|35|^CAR^^14^^^SPC^EndPt|^155.0|||
'L|1|N
'B9
'
sBuf = ""
sBuf = sBuf & ""
sBuf = sBuf & "1H|@^\|URM-bvwh85WA-40||GeneXpert PC^GeneXpert^4.4a|||||1||P|1394-97|20191118122055" & vbCr
sBuf = sBuf & "P|1|||1" & vbCr
sBuf = sBuf & "O|1|1||^^^C|R|20191107163408|||||||||ORH||||||||||F" & vbCr
sBuf = sBuf & "R|1|^C^^diff^Xpert C.difficile G3^4^Cdiff^Toxigenic C|Toxigenic C.diff POSITIVE^|||||F||PMC8762|2F5" & vbCr
sBuf = sBuf & "20191107163408|20191107171724|Cepheid-412BW35^815910^655123^53437082^32001^20210110" & vbCrLf
sBuf = sBuf & "R|2|^C^^diff^^^Cdiff^Toxin B|POS^|||" & vbCr
sBuf = sBuf & "R|3|^C^^diff^^^Cdiff^Toxin B|^26.5|||" & vbCr
sBuf = sBuf & "R|4|^C^^diff^^^Cdiff^Toxin B|^219.0|||" & vbCr
sBuf = sBuf & "R|5|^C^^diff^^^Cdiff^Binary Toxin|POS^|||" & vbCr
sBuf = sBuf & "R02" & vbCrLf
sBuf = sBuf & "3|6|^C^^diff^^^Cdiff^Binary Toxin|^26.2|||" & vbCr
sBuf = sBuf & "R|7|^C^^diff^^^Cdiff^Binary Toxin|^742.0|||" & vbCr
sBuf = sBuf & "R|8|^C^^diff^^^Cdiff^TcdC|NEG^|||" & vbCr
sBuf = sBuf & "R|9|^C^^diff^^^Cdiff^TcdC|^0|||" & vbCr
sBuf = sBuf & "R|10|^C^^diff^^^Cdiff^TcdC|^20.0|||" & vbCr
sBuf = sBuf & "R|11|^C^^diff^^^Cdiff^SPC|NA^|||" & vbCr
sBuf = sBuf & "R|12|^C^^diff^^^Cdi4A" & vbCr
sBuf = sBuf & "4ff^SPC|^32.5|||" & vbCrLf
sBuf = sBuf & "R|13|^C^^diff^^^Cdiff^SPC|^338.0|||" & vbCr
sBuf = sBuf & "R|14|^C^^diff^Xpert C.difficile G3^4^027^027|027 PRESUMPTIVE NEG^|||||F||PMC8762|20191107163408|20191107171724|Cepheid-412BW35^815910^655123^53437082^32001^20210110" & vbCr
sBuf = sBuf & "R|15|^C^^diff^^^027^Tox46" & vbCr
sBuf = sBuf & "5in B|POS^|||" & vbCrLf
sBuf = sBuf & "R|16|^C^^diff^^^027^Toxin B|^26.5|||" & vbCr
sBuf = sBuf & "R|17|^C^^diff^^^027^Toxin B|^219.0|||" & vbCr
sBuf = sBuf & "R|18|^C^^diff^^^027^Binary Toxin|POS^|||" & vbCr
sBuf = sBuf & "R|19|^C^^diff^^^027^Binary Toxin|^26.2|||" & vbCr
sBuf = sBuf & "R|20|^C^^diff^^^027^Binary Toxin|^742.0|||" & vbCr
sBuf = sBuf & "R|21|^C^^diff^^^027^TcdC|N35" & vbCr
sBuf = sBuf & "6EG^|||" & vbCrLf
sBuf = sBuf & "R|22|^C^^diff^^^027^TcdC|^0|||" & vbCr
sBuf = sBuf & "R|23|^C^^diff^^^027^TcdC|^20.0|||" & vbCr
sBuf = sBuf & "R|24|^C^^diff^^^027^SPC|NA^|||" & vbCr
sBuf = sBuf & "R|25|^C^^diff^^^027^SPC|^32.5|||" & vbCr
sBuf = sBuf & "R|26|^C^^diff^^^027^SPC|^338.0|||" & vbCr
sBuf = sBuf & "L|1|N" & vbCr
sBuf = sBuf & "7C" & vbCrLf
sBuf = sBuf & ""

'sBuf = sBuf & "L|1|N" & vbCr
'sBuf = sBuf & "7C" & vbCrLf
'sBuf = sBuf & ""

sBuf = ""
sBuf = sBuf & "1H|@^\|URM-ea4D95WA-94||GeneXpert PC^GeneXpert^4.4a|||||1||P|1394-97|20191118144959" & vbCr
sBuf = sBuf & "P|1|||190021535633" & vbCr
sBuf = sBuf & "O|1|190021535633||^^^CAR|R|20191118094236|||||||||ORH||||||||||F" & vbCr
sBuf = sBuf & "R|1|^CAR^^11^Xpert Carba-R^2^IMP1^|IMP1 NOT DETECTED^|||||F||PMC8762|20198B" & vbCr
sBuf = sBuf & "21118094236|20191118103318|Cepheid-412BW35^814630^691845^515886741^08403^20210321" & vbCrLf
sBuf = sBuf & "R|2|^CAR^^11^^^IMP1^|NEG^|||" & vbCr
sBuf = sBuf & "R|3|^CAR^^11^^^IMP1^Ct|^0|||" & vbCr
sBuf = sBuf & "R|4|^CAR^^11^^^IMP1^EndPt|^2.0|||" & vbCr
sBuf = sBuf & "R|5|^CAR^^11^^^SPC^|PASS^|||" & vbCr
sBuf = sBuf & "R|6|^CAR^^11^^^SPC^Ct|^33.8|||" & vbCr
sBuf = sBuf & "R|7|^CACC" & vbCrLf
sBuf = sBuf & "3R^^11^^^SPC^EndPt|^63.0|||" & vbCr
sBuf = sBuf & "R|8|^CAR^^15^Xpert Carba-R^2^VIM^|VIM NOT DETECTED^|||||F||PMC8762|20191118094236|20191118103318|Cepheid-412BW35^814630^691845^515886741^08403^20210321" & vbCr
sBuf = sBuf & "R|9|^CAR^^15^^^VIM^|NEG^|||" & vbCr
sBuf = sBuf & "R|10|^CAR^^15^^^VIM^Ct|^0|||" & vbCr
sBuf = sBuf & "R|1133" & vbCr
sBuf = sBuf & "4|^CAR^^15^^^VIM^EndPt|^3.0|||" & vbCrLf
sBuf = sBuf & "R|12|^CAR^^15^^^SPC^|PASS^|||" & vbCr
sBuf = sBuf & "R|13|^CAR^^15^^^SPC^Ct|^33.8|||" & vbCr
sBuf = sBuf & "R|14|^CAR^^15^^^SPC^EndPt|^63.0|||" & vbCr
sBuf = sBuf & "R|15|^CAR^^13^Xpert Carba-R^2^NDM^|NDM NOT DETECTED^|||||F||PMC8762|20191118094236|20191118103318|Cepheid-412BW3501" & vbCr
sBuf = sBuf & "5^814630^691845^515886741^08403^20210321" & vbCrLf
sBuf = sBuf & "R|16|^CAR^^13^^^NDM^|NEG^|||" & vbCr
sBuf = sBuf & "R|17|^CAR^^13^^^NDM^Ct|^0|||" & vbCr
sBuf = sBuf & "R|18|^CAR^^13^^^NDM^EndPt|^1.0|||" & vbCr
sBuf = sBuf & "R|19|^CAR^^13^^^SPC^|PASS^|||" & vbCr
sBuf = sBuf & "R|20|^CAR^^13^^^SPC^Ct|^33.8|||" & vbCr
sBuf = sBuf & "R|21|^CAR^^13^^^SPC^EndPt|^63.0|||" & vbCr
sBuf = sBuf & "R|22|^CAR^^8E" & vbCr
sBuf = sBuf & "612^Xpert Carba-R^2^KPC^|KPC NOT DETECTED^|||||F||PMC8762|20191118094236|20191118103318|Cepheid-412BW35^814630^691845^515886741^08403^20210321" & vbCrLf
sBuf = sBuf & "R|23|^CAR^^12^^^KPC^|NEG^|||" & vbCr
sBuf = sBuf & "R|24|^CAR^^12^^^KPC^Ct|^0|||" & vbCr
sBuf = sBuf & "R|25|^CAR^^12^^^KPC^EndPt|^2.0|||" & vbCr
sBuf = sBuf & "R|26|^E9" & vbCrLf
sBuf = sBuf & "7CAR^^12^^^SPC^|PASS^|||" & vbCrLf
sBuf = sBuf & "R|27|^CAR^^12^^^SPC^Ct|^33.8|||" & vbCr
sBuf = sBuf & "R|28|^CAR^^12^^^SPC^EndPt|^63.0|||" & vbCr
sBuf = sBuf & "R|29|^CAR^^14^Xpert Carba-R^2^OXA48^|OXA48 NOT DETECTED^|||||F||PMC8762|20191118094236|20191118103318|Cepheid-412BW35^814630^691845^515886741^08403^248" & vbCrLf
sBuf = sBuf & "00210321" & vbCr
sBuf = sBuf & "R|30|^CAR^^14^^^OXA48^|NEG^|||" & vbCr
sBuf = sBuf & "R|31|^CAR^^14^^^OXA48^Ct|^0|||" & vbCr
sBuf = sBuf & "R|32|^CAR^^14^^^OXA48^EndPt|^6.0|||" & vbCr
sBuf = sBuf & "R|33|^CAR^^14^^^SPC^|PASS^|||" & vbCr
sBuf = sBuf & "R|34|^CAR^^14^^^SPC^Ct|^33.8|||" & vbCr
sBuf = sBuf & "R|35|^CAR^^14^^^SPC^EndPt|^63.0|||" & vbCr
sBuf = sBuf & "L|1|N" & vbCr
sBuf = sBuf & "B0" & vbCrLf
sBuf = sBuf & "" & vbCr


    Call RcvSocketData(sBuf)
    
End Sub

Private Sub Form_Activate()
    MainFrm.lblMenuNm = Me.Caption
    Me.MDIActiveX.WindowState = ccMaximize
End Sub

Private Sub Form_Load()
    Me.Caption = mEqpKey
    
    Me.MousePointer = vbHourglass
    
    Set mIntErrors = New clsIISIntErrors
    Set mIntLib = New clsIISInterface
    Set mOrder = New clsIISIntOrder
    
    Call CtlClear
    Call mIntLib.SetConfig(mEqpCd, mEqpKey)
    'Call GetEqpComm
    'Call GetWinSockComm
    
    DoEvents
    
'    TxtIP = Winsock1.LocalIP
    Winsock1.LocalPort = CInt(10004)
    Winsock1.Listen
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Deactivate()
    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mOrder = Nothing
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISGeneXpert = Nothing
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
    Dim strQcFg As String   'QC유무
    Dim strInfo As String   '수신한 추가정보
    Dim strTemp As String
    Dim i       As Long

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

            '## 1.0.2: 이상대(2005-06-20)
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
            tblResult.Col = TResultEnum.ccLISResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
            tblResult.Col = TResultEnum.ccInfo:
                strInfo = mGetP(strTemp, TResultEnum.ccInfo, DIV)
                If Trim$(strInfo) <> "" Then
                    tblResult.Text = vbCrLf & Space(2) & strInfo & vbCrLf
                End If
            tblResult.Col = TResultEnum.ccEqpResult
                tblResult.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
                If Trim$(strInfo) = "" Then
                    tblResult.ForeColor = vbBlack
                Else
                    tblResult.ForeColor = vbRed
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
            Call mIntLib.WriteLog(Buffer, ccEqp)

            lngBufLen = Len(Buffer)
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)

                Select Case mIntLib.Phase
                    Case 1      '## Estabilshment Phase
                        Select Case BufChar
                            Case ENQ
                                Call mIntLib.ClearBuffer
                                mIntLib.Phase = 2
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                            Case ACK
                                If mIntLib.state = "Q" Then Call SendOrder
                        End Select
                    Case 2      '## Transfer Phase
                        Select Case BufChar
                            Case ENQ
                                Call mIntLib.ClearBuffer
                                MSComm.Output = ACK
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
                            Case vbCr, vbLf
                            Case EOT
                                mIntLib.Phase = 1
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
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                        End Select
                    Case 4      '## Termination Phase
                        Select Case BufChar
                            Case STX
                                mIntLib.Phase = 2
                            Case EOT
                                Call EditRcvData
                                If mIntLib.state = "Q" Then
                                    mIntLib.SndPhase = 1
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
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeq       As String   '수신한 Sequence
    Dim strRackNo    As String   '수신한 Rack Or Disk No
    Dim strTubePos   As String   '수신한 Tube Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strIntName   As String   '수신한 장비기준 검사명
    Dim strResult    As String   '수신한 결과
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim strTemp3     As String
    Dim varRcv       As Variant
    Dim intCnt       As Integer
    Dim blnMTB       As Boolean
    Dim blnRIF       As Boolean
    Dim blnSPC       As Boolean
    Dim strFName     As String
    
    Dim strTemp      As String
    Dim strCtNm      As String
    Dim strCtVal     As String
    
    blnMTB = False
    blnRIF = False
    blnSPC = False
    
    Set objIntNms = mIntLib.IntNms
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        
            Debug.Print strRcvBuf
        '-- 2016.12.19 추가
        Call mIntLib.WriteLog(strRcvBuf & vbCrLf, ccEqp)
        
        varRcv = Split(strRcvBuf, vbCr)
        For intCnt = 0 To UBound(varRcv)
            strRcvBuf = varRcv(intCnt)
            strType = Mid$(strRcvBuf, 1, 1)
            Select Case strType
                Case "H"    '## Header
                Case "Q"    '## Request Information
                Case "P"    '## Patient
                    strBarNo = Format$(mGetP(strRcvBuf, 5, "|"), String$(SPCLEN, "#"))
                    '2P|1|150005776650|||||||||||||||||||||||49
    
                    Set objIntInfo = New clsIISIntInfo
                    With objIntInfo
                        .BarNo = strBarNo
                        '.SpcPos = strTubePos & "/" & strRackNo
                    End With
                Case "O"    '## Order
                    '3O|1|||Flu A+B|||||||||||P4C
                Case "R"    '## Result
                    '## 장비기준 검사명, 결과, Abnormal Flag

                    '-- 2016.12.19 수정
                    strTemp = mGetP(strRcvBuf, 3, "|")
                    strCtVal = mGetP(mGetP(strRcvBuf, 4, "|"), 2, "^") 'Ct값
                    
                    'R|1|^CAR^^11  ^Xpert Carba-R       ^2^IMP1 ^|IMP1 NOT DETECTED^|||||F||PMC8762|20192F
'코로나 전
'R|1 |^C        ^^diff  ^Xpert C.difficile G3^4^Cdiff       ^Toxigenic C    |Toxigenic C.diff POSITIVE  ^|||||F||PMC8762|2A8
'R|14|^C        ^^diff  ^Xpert C.difficile G3^4^027         ^027            |027 PRESUMPTIVE NEG        ^|||||F||PMC8762|20191107163408|20191107171724|Cepheid-412BW35^815910^655123^53437082^32001^20210110
                    
'코로나 이후
'R|1 |^C,diff   ^^Cdiff ^Xpert C.difficile G3^4^Toxigenic C ^diff           |NEGATIVE                   ^|||||F||<None>|20200915110548|20200915114851|Cepheid-412BW35^815910^602496^57451797^32415^20210912
'2021-03-03 변경
'R|1 |^C,diff   ^^Cdiff ^Xpert C.difficile G3^4^Toxigenic C.diff^|POSITIVE^|||||F||<None>|20210303170108|20210303174423|Cepheid-0F09668^740467^896757^60290436^32801^20220403|

'R|14|^C,diff   ^^027   ^Xpert C.difficile G3^4^027         ^               |PRESUMPTIVE NEG            ^|||||F||<None>|20200915110548|20200915114851|Cepheid-412BW35^815910^602496^57451797^32415^20210912
'2021-03-03 변경
'R|14|^C,diff   ^^027   ^Xpert C.difficile G3^4^027^|PRESUMPTIVE NEG^|||||F||<None>|20210303170108|20210303174423|Cepheid-0F09668^740467^896757^60290436^32801^20220403|
                    
                    
'R|1|^C,diff^^Cdiff^Xpert C.difficile G3^4^Toxigenic C^diff|NEGATIVE^|||||F||<None>|20200915110548|20200915114851|Cepheid-412BW35^815910^602496^57451797^32415^20210912
'R|14|^C,diff^^027^Xpert C.difficile G3^4^027^|PRESUMPTIVE NEG^|||||F||<None>|20200915110548|20200915114851|Cepheid-412BW35^815910^602496^57451797^32415^20210912
                    
'R|1|^^^COVID19^Xpert Xpress SARS-CoV-2^2^^|POSITIVE^|||||F||123456|20201016111033|20201016115625|Cepheid-412BW35^815910^724738^386728066^03307^20210919
'R|2|^^^COVID19^^^E^|POS^|||
'R|3|^^^COVID19^^^E^Ct|^31.1|||
'R|4|^^^COVID19^^^E^EndPt|^408.0|||
'R|5|^^^COVID19^^^N2^|POS^|||
'R|6|^^^COVID19^^^N2^Ct|^388
                    
                    
                    '^^^COVID19^Xpert Xpress SARS-CoV-2^2^^
                    strIntBase = mGetP(strTemp, 4, "^") '숫자채널
                    strIntName = mGetP(strTemp, 7, "^") '검사명
                    strCtNm = mGetP(strTemp, 8, "^") 'Ct구분
                    
                    
                    
                    If UCase(strIntBase) = "DIFF" Then
                        If UCase(strIntName) = "CDIFF" Or UCase(strIntName) = "027" Then
                            strIntBase = strIntBase & strIntName
                            strIntBase = UCase(strIntBase)
                        End If
                    End If
                    
                    '-- 2020-09-15 추가
                    If UCase(strIntBase) = "CDIFF" Or UCase(strIntBase) = "027" Then
                        If UCase(strIntName) = "TOXIGENIC C" Or UCase(strIntName) = "TOXIGENIC C.DIFF" Then
                            strIntBase = strIntBase & "TOXIC"
                            strIntBase = UCase(strIntBase)
                        ElseIf UCase(strIntName) = "027" Then
                            strIntBase = strIntBase & strIntName
                            strIntBase = UCase(strIntBase)
                        End If
                    End If
                    
                    strFName = ""
                    strResult = ""

                    '-- MTB
                    If strIntBase = "1" And UCase(strIntName) = "MTB" Then
                        strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
'                        If InStr(strResult, "LOW") > 0 Then
'                            strResult = "DETECTED LOW"
'                        ElseIf InStr(strResult, "HIGH") > 0 Then
'                            strTemp2 = "DETECTED HIGH"
'                        ElseIf InStr(strResult, "NOT") > 0 Then
'                            strResult = "NOT DETECTED"
'                        End If
                    End If
                    
                    '-- RIF
                    If strIntBase = "3" And UCase(strIntName) = "RIF RESISTANCE" Then
                        strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
'                        If strResult = "" Then
'                            strResult = "NA"
'                        ElseIf InStr(strResult, "NOT") > 0 Then
'                            strResult = "NOT DETECTED"
'                        Else
'                            strResult = "DETECTED"
'                        End If

                        If strResult = "" Then
                            strResult = "NA"
                        End If
                    
                    End If
                    
                    '-- 2016.12.19 추가
                    '==> 경우의 수를 봐야 하기 때문에 일단은 장비에서 나온데로 보여줌
                    '-- IMP1
                    If strIntBase = "11" And UCase(strIntName) = "IMP1" Then
                        strFName = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                        If strFName = "Xpert Carba-R" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            '실제결과 "IMP1 NOT DETECTED"로 나옴
                            strResult = Replace(strResult, "IMP1 ", "")
                        End If
                    End If
                    
                    '-- VIM
                    If strIntBase = "15" And UCase(strIntName) = "VIM" Then
                        strFName = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                        If strFName = "Xpert Carba-R" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            '실제결과 "VIM NOT DETECTED"로 나옴
                            strResult = Replace(strResult, "VIM ", "")
                        End If
                    End If
                    
                    '-- NDM
                    If strIntBase = "13" And UCase(strIntName) = "NDM" Then
                        strFName = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                        If strFName = "Xpert Carba-R" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            '실제결과 "NDM NOT DETECTED"로 나옴
                            strResult = Replace(strResult, "NDM ", "")
                        End If
                    End If
                    
                    '-- KPC
                    If strIntBase = "12" And UCase(strIntName) = "KPC" Then
                        strFName = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                        If strFName = "Xpert Carba-R" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            '실제결과 "KPC NOT DETECTED"로 나옴
                            strResult = Replace(strResult, "KPC ", "")
                        End If
                    End If
                    
                    '-- OXA48
                    If strIntBase = "14" And UCase(strIntName) = "OXA48" Then
                        strFName = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                        If strFName = "Xpert Carba-R" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            '실제결과 "OXA48 NOT DETECTED"로 나옴
                            strResult = Replace(strResult, "OXA48 ", "")
                        End If
                    End If
                    
'R|1|^CAR^^11^Xpert Carba-R^2^IMP1^|IMP1 NOT DETECTED^|||||F||PMC8762|20192F
'R|1|^C^^diff^Xpert C.difficile G3^4^Cdiff^Toxigenic C|Toxigenic C.diff POSITIVE^|||||F||PMC8762|2A8
'R|14|^C^^diff^Xpert C.difficile G3^4^027^027|027 PRESUMPTIVE NEG^|||||F||PMC8762|20191107163408|20191107171724|Cepheid-412BW35^815910^655123^53437082^32001^20210110
                    
'R|1|^C^^diff^Xpert C.difficile G3^4^Cdiff^Toxigenic C|Toxigenic C.diff POSITIVE^|||||F||PMC8762|2F5
'diffCdiff
                    
'코로나 전
'R|1 |^C        ^^diff  ^Xpert C.difficile G3^4^Cdiff       ^Toxigenic C    |Toxigenic C.diff POSITIVE  ^|||||F||PMC8762|2A8
'R|14|^C        ^^diff  ^Xpert C.difficile G3^4^027         ^027            |027 PRESUMPTIVE NEG        ^|||||F||PMC8762|20191107163408|20191107171724|Cepheid-412BW35^815910^655123^53437082^32001^20210110
                    
'코로나 이후
'R|1 |^C,diff   ^^Cdiff ^Xpert C.difficile G3^4^Toxigenic C ^diff           |NEGATIVE                   ^|||||F||<None>|20200915110548|20200915114851|Cepheid-412BW35^815910^602496^57451797^32415^20210912
'R|14|^C,diff   ^^027   ^Xpert C.difficile G3^4^027         ^               |PRESUMPTIVE NEG            ^|||||F||<None>|20200915110548|20200915114851|Cepheid-412BW35^815910^602496^57451797^32415^20210912
                    
                    
                    '-- CDIFF
                    If UCase(strIntBase) = "DIFFCDIFF" And UCase(strIntName) = "CDIFF" Then
                        strFName = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                        If strFName = "Xpert C.difficile G3" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            strResult = Replace(strResult, "Toxigenic C.diff ", "")
                        End If
                    End If

                    '-- 2020-09-15 수정 CDIFF
                    'If UCase(strIntBase) = "CDIFFTOXIC" And UCase(strIntName) = "TOXIGENIC C" Then
                    '-- 2021-03-03 수정
                    If UCase(strIntBase) = "CDIFFTOXIC" And UCase(strIntName) = "TOXIGENIC C.DIFF" Then
                        strFName = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                        If strFName = "Xpert C.difficile G3" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            strResult = Replace(strResult, "Toxigenic C.diff ", "")
                        End If
                    End If

                    '-- 027 '027 PRESUMPTIVE NEG^
                    If UCase(strIntBase) = "DIFF027" And UCase(strIntName) = "027" Then
                        strFName = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                        If strFName = "Xpert C.difficile G3" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            '실제결과 "OXA48 NOT DETECTED"로 나옴
                            strResult = Replace(strResult, "027 PRESUMPTIVE ", "")
                            If UCase(strResult) = "NEG" Then
                                strResult = "NEGATIVE"
                            End If
                            If UCase(strResult) = "POS" Then
                                strResult = "POSITIVE"
                            End If
                            
                        End If
                    End If
                   
'R|14|^C,diff^^027^Xpert C.difficile G3^4^027^|PRESUMPTIVE NEG^|||||F||<None>|20200915111419|20200915115802|Cepheid-412BW35^814630^691845^57451776^32415^20210912
                   
                   
                    '-- 2020-09-15 수정 027 '027 PRESUMPTIVE NEG^
                    If UCase(strIntBase) = "027027" And UCase(strIntName) = "027" Then
                        strFName = mGetP(mGetP(strRcvBuf, 3, "|"), 5, "^")
                        If strFName = "Xpert C.difficile G3" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            '실제결과 "OXA48 NOT DETECTED"로 나옴
                            strResult = Replace(strResult, "PRESUMPTIVE ", "")
                            If UCase(strResult) = "NEG" Then
                                strResult = "NEGATIVE"
                            End If
                            If UCase(strResult) = "POS" Then
                                strResult = "POSITIVE"
                            End If
                            
                        End If
                    End If
                   
                    If UCase(strIntBase) = "COVID19" Then
                        If mGetP(strRcvBuf, 2, "|") = "1" Then
                            strResult = mGetP(mGetP(strRcvBuf, 4, "|"), 1, "^")
                            If InStr(UCase(strResult), "NEG") > 0 Then
                                strResult = "NEGATIVE"
                            End If
                            If InStr(UCase(strResult), "POS") > 0 Then
                                strResult = "POSITIVE"
                            End If
                        End If
                    End If
                    If InStr(strResult, "INVA") > 0 Then
                        strResult = ""
                    End If
                    
                    If strResult <> "" Then
                        If objIntNms.ExistIntBase(strIntBase) Then
                            Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                                 strResult, strResult)
                        
                            mIntLib.state = "R"
                        
                        End If
                    End If
                    strResult = ""
                    
                    If UCase(strCtNm) = "CT" Then
                        'If strIntName = "KPC" Or strIntName = "NDM" Or strIntName = "VIM" Or strIntName = "IMP1" Or strIntName = "OXA48" Or strIntName = "SPC" Then
                        If strIntName = "KPC" Or strIntName = "NDM" Or strIntName = "VIM" Or strIntName = "IMP1" Or strIntName = "OXA48" Then
                            
                            strIntBase = strIntName
                            If IsNumeric(strCtVal) Then
                                strCtVal = Val(strCtVal)
                            End If
                            strResult = strCtVal
                    
                            If strResult <> "" Then
                                If objIntNms.ExistIntBase(strIntBase) Then
                                    Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                                         strResult, strResult)
                    
                                    mIntLib.state = "R"
                    
                                End If
                            End If
                            strResult = ""
                            strCtVal = ""
                        Else
                            If strIntName = "SPC" And blnSPC = False Then
                                blnSPC = True
                                
                                strIntBase = strIntName
                                strResult = strCtVal
                        
                                If strResult <> "" Then
                                    If objIntNms.ExistIntBase(strIntBase) Then
                                        Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                                             strResult, strResult)
                        
                                        mIntLib.state = "R"
                        
                                    End If
                                End If
                                strResult = ""
                                strCtVal = ""
                                
                            End If
                        End If
                    End If
                    
                Case "L"    '## Terminator
                    '## DB에 결과저장
                    If mIntLib.state = "R" Then
                        Call SaveServer(objIntInfo)
                        Set objIntInfo = Nothing
                        mIntLib.state = ""
                    End If
            End Select
        Next
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

        Call SetComplete2(objAccInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)

        '## ClientDb, Server에 결과저장
        Call mIntLib.SaveClientDb(strSpcYy, lngSpcNo)
        Call mIntLib.SaveResult(strSpcYy, lngSpcNo)
        
        Call mIntLib.Remove(strSpcYy, lngSpcNo)
        Set objAccInfo = Nothing
        
        StatusBar.Panels(2).Text = "검체번호:" & strBarNo & " 를 정상적으로 결과저장 했습니다."
    End If

    '## tblReady에서 전송된 검체삭제
    With tblReady
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
            'If CStr(vBarNo) = strBarNo Then
            If Mid(CStr(vBarNo), 1, 11) = Mid(strBarNo, 1, 11) Then
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
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput As String     '송신할 데이터
    
    Select Case mIntLib.SndPhase
        Case 1  '## Header
            strOutput = mIntLib.GetFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            mIntLib.SndPhase = 2
            
        Case 2  '## Patient
            strOutput = mIntLib.GetFrameNo & "P|1" & vbCr & ETX
            mIntLib.SndPhase = 4
            
        Case 3  '## No Order
            
        Case 4  '## Order
            With mOrder
                If .NoOrder = True Then
                    '## 접수정보가 없을경우
                    strOutput = mIntLib.GetFrameNo & "O|1|" & .BarNo & "|" & .Seq & "^" & .RackNo & _
                                "^" & .TubePos & "^^SAMPLE^NORMAL|ALL" & _
                                "|R||||||C||||||||||||||Q" & vbCr & ETX
                    mIntLib.SndPhase = 5
                Else
                    If .IsSending = False Then  '## 최초 보낼때
                        strOutput = "O|1|" & .BarNo & "|" & .Seq & "^" & .RackNo & "^" & .TubePos & _
                                    "^^SAMPLE^NORMAL|" & .GetOrder & "|R||||||N||||||||||||||Q"
                        If Len(strOutput) > 230 Then
                            .IsSending = True
                            .Order = Mid$(strOutput, 231)
                            strOutput = mIntLib.GetFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                            mIntLib.SndPhase = 4
                        Else
                            strOutput = mIntLib.GetFrameNo & strOutput & vbCr & ETX
                            mIntLib.SndPhase = 5
                        End If
                    Else                        '## 남은 문자열이 있을때
                        strOutput = .Order
                        If Len(strOutput) > 230 Then
                            .Order = Mid$(strOutput, 231)
                            strOutput = mIntLib.GetFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                            mIntLib.SndPhase = 4
                        Else
                            .IsSending = False
                            strOutput = mIntLib.GetFrameNo & strOutput & vbCr & ETX
                            mIntLib.SndPhase = 5
                        End If
                    End If
                End If
            End With
            
        Case 5  '## Termianator
            strOutput = mIntLib.GetFrameNo & "L|1" & vbCr & ETX
            mIntLib.SndPhase = 6
            
        Case 6  '## EOT
            mIntLib.state = ""
            MSComm.Output = EOT
            Call mIntLib.WriteLog(EOT, ccPCLog)
            Exit Sub
    End Select
    
    strOutput = STX & strOutput & mOrder.GetChkSum(strOutput) & vbCrLf
    MSComm.Output = strOutput
    Call mIntLib.WriteLog(strOutput, ccPCLog)
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblReady에 정보표시
'   인수 :
'       - pAccInfo : 접수내역 클래스
'-----------------------------------------------------------------------------'
Private Sub SetReady(ByVal pAccInfo As clsIISAccInfo)
    Dim lngWorkNo As Long   'WorkNo

    With tblReady
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If

        .Col = TReadyEnum.ccNo:     .Text = mOrder.TubePos & "/" & Mid$(mOrder.RackNo, 2)
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
    Dim vBarNo      As Variant          'Spread의 바코드번호
    Dim vTubePos    As Variant          'Spread의 Tube Position
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

        i = 0
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
                .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                        DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & _
                        DIV & mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal) & DIV & _
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
'   기능 : 소켓포트 Open
'-----------------------------------------------------------------------------'
Private Sub GetWinSockComm()

    If Winsock1.state <> sckClosed Then
        Winsock1.Close
    End If
      
    Winsock1.LocalPort = 10002
    Winsock1.Listen

'    lblWinSock.Caption = "Socket : Connecting.."

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
            Error.SetLog App.EXEName, "frmIISElecsys2010", "GetEqpComm", strErrMsg, Now
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
    If err.Number = 8005 Then
        strErrMsg = mEqpCd & " 장비에 설정된 포트가 이미 사용중입니다."
        Error.SetLog App.EXEName, "frmIISElecsys2010", "GetEqpComm", strErrMsg, Now
        Call mIntLib_EqpError("E005")
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 수신한 Abnormal Flag에 대한 설명조회
'-----------------------------------------------------------------------------'
Private Function GetContent(ByVal pFlags As String) As String
    Dim strFlag     As String   'Abnormal Flag
    Dim strContent  As String   'Abnormal Content
    Dim strTemp     As String
    Dim i           As Long
    
    If pFlags = "" Then Exit Function
    
    For i = 1 To Len(Trim(pFlags))
        strFlag = Mid$(pFlags, i, 1)
        If i = 1 Then
            strTemp = Space(2) & "[Abnormality flag]: " & strFlag & vbCrLf
        Else
            strTemp = strTemp & vbCrLf & Space(2) & "[Abnormality flag]: " & strFlag & vbCrLf
        End If
        
        strContent = ""
        Select Case strFlag
            Case "/": strContent = "Test no performed: test has been requisitioned but not performed due to any reason."
            Case "S": strContent = "Result extracted for repeat run"
            Case "?": strContent = "Calculation unable due to abnormal photometric data. UNIT in STOP mode (Incl. Lamp OFF), etc."
            Case "n": strContent = "8087 error"
            Case "R": strContent = "Reagent level detection error"
            Case "#": strContent = "Sample level detection error"
            Case "!": strContent = "A/D error of photometry"
            Case ">": strContent = "The absolute OD value is over 2.665."
            Case "<": strContent = "The absolute OD value is under 0.99."
            Case "-": strContent = "The final result is negative."
            Case "U": strContent = "Reagent absorbance value at P0 of Reagent Blank run, is smaller than the lower limit of the Parameter."
            Case "u": strContent = "Reagent absorbance value at P0 or p8 is lower than the lower limit specified in the Parameters in routine run."
            Case "Y": strContent = "Reagent absorbance value at P16 of Reagent Blank run, is greater than the upper limit of the Parameter."
            Case "y": strContent = "Reagent absorbance value at P0 or p8 is higher than the upper limit specified in the Parameters in routine run."
            Case "@": strContent = "Abnormally high result: absorbance of every wavelength is more than 2.5."
            Case "$": strContent = "No linearity validation conducted because less than 3 data obtained in the kinetics."
            Case "D": strContent = "Too quick reaction slope in increasing kinetics, absorbance at P-START is higher than MAX. OD in increasing FIXED assay, or too slow reaction slope in decreasing kinetics (=no reaction observed)"
            Case "B": strContent = "Too quick reaction slope in increasing kinetics, or absorbance at P-END is lower than MIN. OD in increasing FIXED assay."
            Case "*": strContent = "Linearity error in kinetics"
            Case "P": strContent = "Result higher than DECIDE RANGE designated in parameters."
            Case "N": strContent = "Result lower than DECIDE RANGE designated in parameters."
            Case "&": strContent = "Data check 2 error"
            Case "Z": strContent = "Data check 1 error"
            Case "F": strContent = "Result higher than the dynamic range specified in the Parameters"
            Case "G": strContent = "Result lower than the dynamic range specified in the Parameters"
            Case "p": strContent = "Result beyond the panic value specified in the Parameters"
            Case "T": strContent = "Abnormality found in the Inter-Item Check"
            Case "H": strContent = "Result higher than the normal value range specified in the Parameters"
            Case "L": strContent = "Result lower than the normal value range specified in the Parameters"
            Case "W": strContent = "Abnormality in WB data. Photocal has not been performed."
            Case "J": strContent = "Result higher than the repeat run range specified in the Parameters"
            Case "K": strContent = "Result higher than the repeat run range specified in the Parameters"
        End Select
        
        If strContent <> "" Then
            strTemp = strTemp & Space(2) & "[Content]: " & strContent & vbCrLf
        End If
    Next i
    
    GetContent = vbCrLf & strTemp
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
                End If
            End With
        Case DELETEALL  '## Delete All
            Call mIntLib.AccInfos.RemoveAll
            Call mTblClear(tblReady)
    End Select
End Sub

'Private Sub Winsock1_Close()
'
'    lblWinSock.Caption = "Socket : Close!!!"
'    'Call Sleep(1000)
'    Call GetWinSockComm
'
'End Sub
'
'Private Sub Winsock1_Connect()
'    lblWinSock.Caption = "Socket : Connection OK!"
'End Sub
'
'Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'    Winsock1.Close
'    Winsock1.Accept requestID
'End Sub
'
Public Sub RcvSocketData(ByVal lsData As String)
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    Buffer = lsData
    Call mIntLib.WriteLog(Buffer, ccEqp)

    lngBufLen = Len(Buffer)
    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)

        Select Case mIntLib.Phase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Call mIntLib.ClearBuffer
                        mIntLib.Phase = 2
                        'Call Winsock1.SendData(ACK)
                        sck.ProcSendMessage ACK
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                    Case ACK
                        'If mIntLib.State = "Q" Then Call SendOrder
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case ENQ
                        Call mIntLib.ClearBuffer
                        sck.ProcSendMessage ACK
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
                    Case vbCr
                        mIntLib.BufCnt = mIntLib.BufCnt + 1
                        'mIntLib.Phase = 3
                    Case vbLf
                    Case EOT
                        mIntLib.Phase = 1
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
                        'Call mIntLib.AddBuffer(BufChar)
                        'mIntLib.Phase = IIf(mIntLib.IsETB = False, 4, 2)
                        mIntLib.Phase = 4
                        sck.ProcSendMessage ACK
                        Call mIntLib.WriteLog(ACK, ccPCLog)
                End Select
            Case 4      '## Termination Phase
                Select Case BufChar
                    Case STX
                        mIntLib.Phase = 2
                    Case EOT
                        Call EditRcvData
                        'If mIntLib.State = "Q" Then
                        '    mIntLib.SndPhase = 1
                        '    mIntLib.FrameNo = 0
                        '    MSComm.Output = ENQ
                        '    Call mIntLib.WriteLog(ENQ, ccPCLog)
                        'End If
                        mIntLib.Phase = 1
                End Select
        End Select
    Next i

    'Winsock1.SendData lsTmp

    Exit Sub

ErrHandle:
    Resume Next
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    sck.Accept requestID
    
    Winsock1.Close
    Winsock1.Listen
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strRcvBuffer As String
    Dim strSndBuffer As String
   
    Winsock1.GetData strRcvBuffer
    Debug.Print strRcvBuffer


    strSndBuffer = "ORDER"
    Winsock1.SendData (strSndBuffer)

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Number & " >> " & Description
    Winsock1.Close
End Sub


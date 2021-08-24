VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{D74ED2A2-3650-4720-93BC-FDDD8DCBC769}#1.0#0"; "Han2EngOCX.ocx"
Begin VB.Form frmIISAlinity 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Alinity"
   ClientHeight    =   9180
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9180
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtHigh 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   9630
      TabIndex        =   30
      Text            =   "03271218"
      Top             =   8760
      Width           =   1485
   End
   Begin VB.TextBox txtNormal 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   8100
      TabIndex        =   29
      Text            =   "03271218"
      Top             =   8760
      Width           =   1485
   End
   Begin VB.TextBox txtLow 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      Height          =   270
      Left            =   6570
      TabIndex        =   28
      Text            =   "03271218"
      Top             =   8760
      Width           =   1485
   End
   Begin VB.Timer tmrNow 
      Left            =   5910
      Top             =   210
   End
   Begin HAN2ENGOCXLib.Han2EngOCX Han2EngOCX1 
      Height          =   255
      Left            =   12390
      TabIndex        =   27
      Top             =   8310
      Visible         =   0   'False
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin MSWinsockLib.Winsock wSck 
      Left            =   11850
      Top             =   8220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   4710
      TabIndex        =   26
      Top             =   150
      Visible         =   0   'False
      Width           =   615
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
      Left            =   10650
      Top             =   8160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MDIControls.MDIActiveX MDIActiveX 
      Left            =   11310
      Top             =   8190
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
      SpreadDesigner  =   "frmIISAlinity.frx":0000
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
      SpreadDesigner  =   "frmIISAlinity.frx":03FA
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
      SpreadDesigner  =   "frmIISAlinity.frx":0B44
      TextTip         =   2
   End
   Begin VB.Label Label3 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "QC High"
      Height          =   195
      Left            =   9660
      TabIndex        =   33
      Top             =   8520
      Width           =   1425
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "QC Normal"
      Height          =   195
      Left            =   8130
      TabIndex        =   32
      Top             =   8520
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "QC Low"
      Height          =   195
      Left            =   6600
      TabIndex        =   31
      Top             =   8520
      Width           =   1425
   End
End
Attribute VB_Name = "frmIISAlinity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISAlinity.frm
'   작성자  : 오세원
'   내  용  : Alinity 장비폼
'   작성일  : 2019-12-19
'   버  전  :
'   병  원  :
'       1. 예수병원
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

Dim strDump As String
Dim strRackPos  As String

Private gSckPort    As String
Private gEqp1       As String
Private gEqp2       As String

Private gSpcYY_L    As String
Private gSpcYY_N    As String
Private gSpcYY_H    As String
Private gSpcNo_L    As String
Private gSpcNo_N    As String
Private gSpcNo_H    As String

Private gBarno_L    As String
Private gBarno_N    As String
Private gBarno_H    As String
Private gNow        As String

Private gLow        As String
Private gNormal     As String
Private gHigh       As String



Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property


Private Function GetAlinityConfig(ByVal strConfigNm As String) As String
    Dim strFileName As String
    Dim strReturnedString As String
    
    GetAlinityConfig = ""
    
    strFileName = App.Path & "\Alinity.ini"
    
    strReturnedString = String(1024, " ")
    GetPrivateProfileString "ALINITY", strConfigNm, "", strReturnedString, Len(strReturnedString), strFileName
    strReturnedString = Trim(Replace(strReturnedString, Chr(0), Chr(32), 1, -1, vbBinaryCompare))
    GetAlinityConfig = strReturnedString
    
End Function


Private Sub Command1_Click()

    strDump = ""
    strDump = strDump & "1H|\^&|||Alinity^2.20^52500030827^H1P1O1R1C1Q1L1|||||||P|1|20070801102234" & vbNewLine
    strDump = strDump & "07" & vbNewLine
    strDump = strDump & "2P|1||01155107|01155107" & vbNewLine
    strDump = strDump & "DB" & vbNewLine
    strDump = strDump & "3O|1|17000416672|17000416672^J401^1|^^^161^Anti-HCV^UNDILUTED^P|R||||||||||||||||||||F" & vbNewLine
    strDump = strDump & "0B" & vbNewLine
    strDump = strDump & "4R|1|^^^161^Anti-HCV^UNDILUTED^P^48229HN00^04462^^F|0.06|S/CO||||F||FSE^FSE||20070801102234|iSR03698" & vbNewLine
    strDump = strDump & "42" & vbNewLine
    strDump = strDump & "5R|2|^^^161^Anti-HCV^UNDILUTED^P^48229HN00^04462^^I|Nonreactive|||||F||FSE^FSE||20070801102234|iSR03698" & vbNewLine
    strDump = strDump & "ED" & vbNewLine
    strDump = strDump & "6R|3|^^^161^Anti-HCV^UNDILUTED^P^48229HN00^04462^^P|1448|RLU||||F||FSE^FSE||20070801102234|iSR03698" & vbNewLine
    strDump = strDump & "3C" & vbNewLine
    strDump = strDump & "7L|1" & vbNewLine
    strDump = strDump & "40" & vbNewLine
    strDump = strDump & ""

strDump = ""
strDump = strDump & ""
strDump = strDump & STX & "1H|\^&|||Alinity ci-series^2.6^SCM04183|||||||P|LIS2-A2|20200128083305+0900"
strDump = strDump & ETX & "28"

strDump = strDump & STX & "2P|1"
strDump = strDump & ETX & "3F"

strDump = strDump & STX & "3O|1|0327121801|0327121801^K2472^1^1^5|^^^593^CA 125 II^UNDILUTED|R||||||Q||||||||||||||F"
strDump = strDump & ETX & "F6"

strDump = strDump & STX & "4M|1|INV|MCC_LEVEL1|CO||||032712180"
strDump = strDump & ETX & "C4"

strDump = strDump & STX & "5R|1|^^^593^CA 125 II^UNDILUTED^F|14.9|U/mL|10.70000 - 20.00000|||F||Admin^Admin||20200128083305|Ai04026"
strDump = strDump & ETX & "25"

strDump = strDump & STX & "6M|1|INV|593|CA|||20200126115421|09526FP00"
strDump = strDump & ETX & "6D"

strDump = strDump & STX & "7M|2|INV|593-1|SR|02434|20200807||09526FP00"
strDump = strDump & ETX & "C3"

strDump = strDump & STX & "0R|2|^^^593^CA 125 II^UNDILUTED^P|16796|RLU||||F||Admin^Admin||20200128083305|Ai04026"
strDump = strDump & ETX & "AF"

strDump = strDump & STX & "1R|3|^^^593^CA 125 II^UNDILUTED^G|053574d3-00ac-445a-8603-8c603a3432e8|||||F||Admin^Admin||20200128083305|Ai04026"
strDump = strDump & ETX & "1C"

strDump = strDump & STX & "2O|2|0327121801|0327121801^K2472^1^1^5|^^^658^B-hCG^UNDILUTED|R||||||Q||||||||||||||F"
strDump = strDump & ETX & "6B"

strDump = strDump & STX & "3M|1|INV|MCC_LEVEL1|CO||||032712180"
strDump = strDump & ETX & "C3"

strDump = strDump & STX & "4R|1|^^^658^B-hCG^UNDILUTED^F|6.94|mIU/mL|3.64000 - 6.76000|CNTL\1-2s||F||Admin^Admin||20200128082105|Ai04026"
strDump = strDump & ETX & "97"

strDump = strDump & STX & "5M|1|INV|658|CA|||20200117113357|06112UI00"
strDump = strDump & ETX & "70"

strDump = strDump & STX & "6M|2|INV|658-1|SR|03674|20200624||06112UI00"
strDump = strDump & ETX & "C4"

strDump = strDump & STX & "7R|2|^^^658^B-hCG^UNDILUTED^P|1750|RLU||||F||Admin^Admin||20200128082105|Ai04026"
strDump = strDump & ETX & "E8"

strDump = strDump & STX & "0R|3|^^^658^B-hCG^UNDILUTED^G|03dc3f1c-1078-47ca-b381-c68f6360ce1f|||||F||Admin^Admin||20200128082105|Ai04026"
strDump = strDump & ETX & "86"

strDump = strDump & STX & "1O|3|0327121801|0327121801^K2472^1^1^5|^^^10^AFP^UNDILUTED|R||||||Q||||||||||||||F"
strDump = strDump & ETX & "9F"

strDump = strDump & STX & "2M|1|INV|MCC_LEVEL1|CO||||032712180"
strDump = strDump & ETX & "C2"

strDump = strDump & STX & "3R|1|^^^10^AFP^UNDILUTED^F|3.72|ng/mL|3.06000 - 4.59000|||F||Admin^Admin||20200128083005|Ai04026"
strDump = strDump & ETX & "F8"

strDump = strDump & STX & "4M|1|INV|10|CA||20200215144911|20200116144911|05024FN00"
strDump = strDump & ETX & "E3"

strDump = strDump & STX & "5M|2|INV|10-1|SR|01936|20200516||05024FN00"
strDump = strDump & ETX & "77"

strDump = strDump & STX & "6R|2|^^^10^AFP^UNDILUTED^P|7989|RLU||||F||Admin^Admin||20200128083005|Ai04026"
strDump = strDump & ETX & "2F"

strDump = strDump & STX & "7R|3|^^^10^AFP^UNDILUTED^G|4a544e31-b27b-4bbc-a3f5-0375f97b6863|||||F||Admin^Admin||20200128083005|Ai04026"
strDump = strDump & ETX & "98"

strDump = strDump & STX & "0O|4|0327121801|0327121801^K2472^1^1^5|^^^598^CA19-9XR^UNDILUTED|R||||||Q||||||||||||||F"
strDump = strDump & ETX & "0B"

strDump = strDump & STX & "1M|1|INV|MCC_LEVEL1|CO||||032712180"
strDump = strDump & ETX & "C1"

strDump = strDump & STX & "2R|1|^^^598^CA19-9XR^UNDILUTED^F|35.59|U/mL|22.90000 - 42.50000|||F||Admin^Admin||20200128083023|Ai04026"
strDump = strDump & ETX & "7A"

strDump = strDump & STX & "3M|1|INV|598|CA|||20200117110808|07477FP00"
strDump = strDump & ETX & "76"

strDump = strDump & STX & "4M|2|INV|598-1|SR|02658|20200603||07477FP00"
strDump = strDump & ETX & "CA"

strDump = strDump & STX & "5R|2|^^^598^CA19-9XR^UNDILUTED^P|9102|RLU||||F||Admin^Admin||20200128083023|Ai04026"
strDump = strDump & ETX & "85"

strDump = strDump & STX & "6R|3|^^^598^CA19-9XR^UNDILUTED^G|11470564-154f-4926-addc-966cf6b4eebf|||||F||Admin^Admin||20200128083023|Ai04026"
strDump = strDump & ETX & "3C"

strDump = strDump & STX & "7O|5|0327121801|0327121801^K2472^1^1^5|^^^47^CEA^SAMP/CNTRL|R||||||Q||||||||||||||F"
strDump = strDump & ETX & "D8"

strDump = strDump & STX & "0M|1|INV|MCC_LEVEL1|CO||||032712180"
strDump = strDump & ETX & "C0"

strDump = strDump & STX & "1R|1|^^^47^CEA^SAMP/CNTRL^F|2.14|ng/mL|1.73000 - 2.60000|||F||Admin^Admin||20200128083041|Ai04026"
strDump = strDump & ETX & "1A"

strDump = strDump & STX & "2M|1|INV|47|CA|||20200116105017|09289FN00"
strDump = strDump & ETX & "36"

strDump = strDump & STX & "3M|2|INV|47-1|SR|00202|20210304||09289FN00"
strDump = strDump & ETX & "7D"

strDump = strDump & STX & "4R|2|^^^47^CEA^SAMP/CNTRL^P|1630|RLU||||F||Admin^Admin||20200128083041|Ai04026"
strDump = strDump & ETX & "47"

strDump = strDump & STX & "5R|3|^^^47^CEA^SAMP/CNTRL^G|20719510-709d-453a-8078-d48e5c330e96|||||F||Admin^Admin||20200128083041|Ai04026"
strDump = strDump & ETX & "DF"

strDump = strDump & STX & "6O|6|0327121801|0327121801^K2472^1^1^5|^^^355^Total PSA^SAMP/CNTRL|R||||||Q||||||||||||||F"
strDump = strDump & ETX & "49"

strDump = strDump & STX & "7M|1|INV|MCC_LEVEL1|CO||||032712180"
strDump = strDump & ETX & "C7"

strDump = strDump & STX & "0R|1|^^^355^Total PSA^SAMP/CNTRL^F|0.544|ng/mL|0.38200 - 0.70900|||F||Admin^Admin||20200128083059|Ai04026"
strDump = strDump & ETX & "D3"

strDump = strDump & STX & "1M|1|INV|355|CA|||20191216200215|07376FN00"
strDump = strDump & ETX & "68"

strDump = strDump & STX & "2M|2|INV|355-1|SR|01117|20200920||07376FN00"
strDump = strDump & ETX & "B2"

strDump = strDump & STX & "3R|2|^^^355^Total PSA^SAMP/CNTRL^P|31303|RLU||||F||Admin^Admin||20200128083059|Ai04026"
strDump = strDump & ETX & "F0"

strDump = strDump & STX & "4R|3|^^^355^Total PSA^SAMP/CNTRL^G|72362eb0-bfbb-4bab-a0c0-b2f5033b6e06|||||F||Admin^Admin||20200128083059|Ai04026"
strDump = strDump & ETX & "DC"

strDump = strDump & STX & "5O|7|0327121801|0327121801^K2472^1^1^5|^^^593^CA 125 II^UNDILUTED|R||||||Q||||||||||||||F"
strDump = strDump & ETX & "FE"

strDump = strDump & STX & "6M|1|INV|MCC_LEVEL1|CO||||032712180"
strDump = strDump & ETX & "C6"

strDump = strDump & STX & "7R|1|^^^593^CA 125 II^UNDILUTED^F|14.5|U/mL|10.70000 - 20.00000|||F||Admin^Admin||20200128083211|Ai04026"
strDump = strDump & ETX & "1F"

strDump = strDump & STX & "0M|1|INV|593|CA|||20200117124450|08510FP00"
strDump = strDump & ETX & "61"

strDump = strDump & STX & "1M|2|INV|593-1|SR|01296|20200718||08510FP00"
strDump = strDump & ETX & "BB"

strDump = strDump & STX & "2R|2|^^^593^CA 125 II^UNDILUTED^P|18627|RLU||||F||Admin^Admin||20200128083211|Ai04026"
strDump = strDump & ETX & "A8"

strDump = strDump & STX & "3R|3|^^^593^CA 125 II^UNDILUTED^G|0675463c-9834-4f3a-b5ea-482bd35bd292|||||F||Admin^Admin||20200128083211|Ai04026"
strDump = strDump & ETX & "BE"

strDump = strDump & STX & "4O|8|0327121801|0327121801^K2472^1^1^5|^^^68^Ferritin^SAMP/CNTRL|R||||||Q||||||||||||||F"
strDump = strDump & ETX & "55"

strDump = strDump & STX & "5M|1|INV|MCC_LEVEL1|CO||||032712180"
strDump = strDump & ETX & "C5"

strDump = strDump & STX & "6R|1|^^^68^Ferritin^SAMP/CNTRL^F|30.59|ng/mL|20.60000 - 38.30000|||F||Admin^Admin||20200128083229|Ai04026"
strDump = strDump & ETX & "41"

strDump = strDump & STX & "7M|1|INV|68|CA|||20200123095605|09292UI00"
strDump = strDump & ETX & "4B"

strDump = strDump & STX & "0M|2|INV|68-1|SR|00604|20201201||09292UI00"
strDump = strDump & ETX & "83"

strDump = strDump & STX & "1R|2|^^^68^Ferritin^SAMP/CNTRL^P|26698|RLU||||F||Admin^Admin||20200128083229|Ai04026"
strDump = strDump & ETX & "0E"

strDump = strDump & STX & "2R|3|^^^68^Ferritin^SAMP/CNTRL^G|2f7ad518-f4ac-42af-819f-0dea029a8983|||||F||Admin^Admin||20200128083229|Ai04026"
strDump = strDump & ETX & "B0"

strDump = strDump & STX & "3O|9|0327121801|0327121801^K2472^1^1^5|^^^774^Insulin^UNDILUTED|R||||||Q||||||||||||||F"
strDump = strDump & ETX & "F3"

strDump = strDump & STX & "4M|1|INV|MCC_LEVEL1|CO||||032712180"
strDump = strDump & ETX & "C4"

strDump = strDump & STX & "5R|1|^^^774^Insulin^UNDILUTED^F|13.4|uU/mL|10.90000 - 16.30000|||F||Admin^Admin||20200128083247|Ai04026"
strDump = strDump & ETX & "98"

strDump = strDump & STX & "6M|1|INV|774|CA|||20200117113227|10511LP39"
strDump = strDump & ETX & "74"

strDump = strDump & STX & "7M|2|INV|774-1|SR|01819|20200608||10511LP39"
strDump = strDump & ETX & "CD"

strDump = strDump & STX & "0R|2|^^^774^Insulin^UNDILUTED^P|47353|RLU||||F||Admin^Admin||20200128083247|Ai04026"
strDump = strDump & ETX & "A2"

strDump = strDump & STX & "1R|3|^^^774^Insulin^UNDILUTED^G|898888fa-c0fb-40df-8dbf-04744154ad1e|||||F||Admin^Admin||20200128083247|Ai04026"
strDump = strDump & ETX & "53"

strDump = strDump & STX & "2L|1"
strDump = strDump & ETX & "3B"
strDump = strDump & ""


'    Call MSComm_OnComm

Call RcvSocketData(strDump)

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
    'TCP-IP 통신사용함
    'Call GetEqpComm
    
    gSckPort = GetAlinityConfig("TCPPORT")
    gEqp1 = GetAlinityConfig("EQP1")
    gEqp2 = GetAlinityConfig("EQP2")

    gLow = GetAlinityConfig("L")
    gNormal = GetAlinityConfig("N")
    gHigh = GetAlinityConfig("H")
    
    txtLow.Text = gLow
    txtNormal.Text = gNormal
    txtHigh.Text = gHigh

    tmrNow.Interval = 60000
    tmrNow.Enabled = True
    
    If gSckPort <> "" And IsNumeric(gSckPort) Then
        wSck.LocalPort = CInt(gSckPort)
        wSck.Listen
    End If
    
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
    Set frmIISAlinity = Nothing
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
    Dim strSpcYY    As String           '검체연도
    Dim lngSpcNo    As Long             '검체번호

    If Row = 0 Then Exit Sub

    Call CtlClear(ccLabel)
    Call mTblClear(tblResult)
    Call tblReady.GetText(TReadyEnum.ccBarNo, Row, vBarNo)
    If vBarNo = "" Then Exit Sub
    
    strSpcYY = Mid$(vBarNo, 1, SPCYYLEN)
    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
    Set objAccInfo = mIntLib.AccInfos(strSpcYY, lngSpcNo)

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

            '## 1.0.1: 오세원(2005-06-24)
            '   - 결과표시 버그수정
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
            tblResult.Col = TResultEnum.ccEqpResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
        Next i
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
            'Buffer = strDump
            
            Call mIntLib.WriteLog(Buffer, ccEqp)

            Debug.Print Buffer
            
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
                                If mIntLib.State = "Q" Then
                                    Call SendOrder
                                Else    '-- Edit by Sewon,Oh(2006.08.02)
                                    MSComm.Output = ACK
                                End If
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
                                    mIntLib.SndPhase = 1
                                    mIntLib.FrameNo = 0
                                    MSComm.Output = ENQ
                                    Call mIntLib.WriteLog(ENQ, ccPCLog)
                                End If
                                mIntLib.Phase = 1
'                                MSComm.Output = ACK
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

Private Sub EditRcvData()
    Dim objIntInfo   As clsIISIntInfo    '인터페이스 검체정보 클래스
    Dim objIntNms    As clsIISIntNms     '장비별 검사항목 컬렉션 클래스
    Dim objBuffer    As clsIISBuffer     '버퍼클래스

    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 바코드번호
    Dim strSeg       As String   '수신한 Segment
    Dim strPos       As String   '수신한 Position
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strIntResult As String   '수신한 결과
    Dim strFlag      As String   '수신한 정량,정성 구분(F:정량,I:정성)
    Dim strResult    As String   'LIS결과
    Dim strTemp      As String
    Dim strEqpCd     As String
    
    Dim strQCFlag    As String
    Dim strSQL       As String
    Dim AdoRS        As ADODB.Recordset
    Dim strSpcYY     As String
    Dim strSpcNo     As String
    Dim strNow       As String
    
    Set objIntNms = mIntLib.IntNms
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        strType = Mid$(strRcvBuf, 2, 1)

        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "Q"    '## Request Information
                '## 바코드번호 조회
                strBarNo = Trim$(mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^"))

                With mOrder
                    .ClsClear
                    .BarNo = strBarNo
                End With
                Call GetOrder(strBarNo)
                mIntLib.State = "Q"

            Case "O"    '## Order
                strQCFlag = ""
                strTemp = mGetP(strRcvBuf, 4, "|")
                strBarNo = mGetP(strTemp, 1, "^")
                strSeg = mGetP(strTemp, 2, "^")
                strPos = mGetP(strTemp, 3, "^")
                strQCFlag = mGetP(strRcvBuf, 12, "|")
                
                '일반
                '3O|1|100011171208     |100011171208^K2462^1^1^18|^^^10^AFP^UNDILUTED|R||||||||||||||||||||F

                'QC : LOW = 0327121801, N = 0327121802, H = 0327121803 12:Q
                '3O|1|0327121801|0327121801^K2472^1^1^5|^^^593^CA 125 II^UNDILUTED|R||||||Q||||||||||||||F
                '3O|1|HIV3|HIV3^K2679^3^1^17|^^^396^HIV Ag/Ab^UNDILUTED|R||||||Q||||||||||||||F
                
'''                If strQCFlag = "Q" Then
'''                    strNow = Format(Now, "yyyymmdd")
'''
'''                    strSQL = ""
'''                    strSQL = strSQL & ""
'''                    strSQL = strSQL & "SELECT DISTINCT b.SPCYY, b.SPCNO " & vbCrLf
'''                    strSQL = strSQL & "  FROM s2lab026 a, s2lab201 b    " & vbCrLf
'''                    strSQL = strSQL & " WHERE a.lotno = '" & strBarNo & "'" & vbCrLf
'''                    strSQL = strSQL & "   AND a.ACCDT = '" & strNow & "'" & vbCrLf
'''
'''                    If Right(strBarNo, 1) = "1" Then
'''                        strSQL = strSQL & "   AND a.levelcd =  'L'      " & vbCrLf
'''                    ElseIf Right(strBarNo, 1) = "2" Then
'''                        strSQL = strSQL & "   AND a.levelcd =  'N'      " & vbCrLf
'''                    ElseIf Right(strBarNo, 1) = "3" Then
'''                        strSQL = strSQL & "   AND a.levelcd =  'H'      " & vbCrLf
'''                    End If
'''
'''                    strSQL = strSQL & "   AND a.rstcd is null           " & vbCrLf
'''                    strSQL = strSQL & "   AND a.WORKAREA = b.WORKAREA   " & vbCrLf
'''                    strSQL = strSQL & "   AND a.ACCDT    = b.ACCDT      " & vbCrLf
'''                    strSQL = strSQL & "   AND a.ACCSEQ   = b.ACCSEQ     " & vbCrLf
'''
'''                    Set AdoRS = New ADODB.Recordset
'''                    Set AdoRS = DbCon.Execute(strSQL, , adCmdText)
'''                    If Not (AdoRS.BOF Or AdoRS.EOF) Then
'''                        strSpcYY = AdoRS.Fields("SPCYY").Value & ""
'''                        strSpcNo = AdoRS.Fields("SPCNO").Value & ""
'''                        strBarNo = strSpcYY & Format$(strSpcNo, String$(SPCNOLEN, "0"))
'''                    End If
'''                    Set AdoRS = Nothing
'''                End If

                If strBarNo = "" Then
                    Exit Sub
                End If
                
                'HIV3 이런것 때문에 IsNumeric 사용함.
                If IsNumeric(strBarNo) Then
                    If strQCFlag = "Q" Then
                        If Right(strBarNo, 1) = "1" Then
                            strSpcYY = gSpcYY_L
                            strSpcNo = gSpcNo_L
                            strBarNo = strSpcYY & Format$(strSpcNo, String$(SPCNOLEN, "0"))
                        ElseIf Right(strBarNo, 1) = "2" Then
                            strSpcYY = gSpcYY_N
                            strSpcNo = gSpcNo_N
                            strBarNo = strSpcYY & Format$(strSpcNo, String$(SPCNOLEN, "0"))
                        ElseIf Right(strBarNo, 1) = "3" Then
                            strSpcYY = gSpcYY_H
                            strSpcNo = gSpcNo_H
                            strBarNo = strSpcYY & Format$(strSpcNo, String$(SPCNOLEN, "0"))
                        End If
                        
                    End If
                End If
                
                Set objIntInfo = New clsIISIntInfo
                With objIntInfo
                    .BarNo = strBarNo
                    .SpcPos = strPos & "/" & strSeg
                    strRackPos = .SpcPos
                End With
                
            '-- 2019.12.19 오세원
            '-- 결과(R) 채널에 있는 코멘트는 Architect 기준임.
            'R 에서 IGbn:= TokenStr(sData, '|', 14);  장비명 나와요
            '1번장비 : Ai04026
            '2번장비 : Ai04025
            '4R|1|^^^598^CA19-9XR^UNDILUTED^F|6.62|U/mL||||F||Admin^Admin||20191220093826|Ai04026

        
            Case "R"    '## Result
                strTemp = mGetP(strRcvBuf, 3, "|")
                strIntBase = mGetP(strTemp, 4, "^")
                strFlag = mGetP(strTemp, 7, "^")
                strIntResult = mGetP(strRcvBuf, 4, "|")
                
                '2020.01.08 : 장비별로 구분하여 저장기능 사용하지 않음
'                '장비별로 구분하여 저장
'                strEqpCd = mGetP(strRcvBuf, 14, "|")
'
'                If strEqpCd = gEqp1 Then
'                    mEqpCd = "E122" '1번장비 Ai04026
'                    EqpCd = "E122"
'                Else
'                    mEqpCd = "E123" '2번장비 Ai04025
'                    EqpCd = "E123"
'                End If
                
                mEqpCd = "E122" '1번장비 Ai04026
                EqpCd = "E122"
                
                Select Case strFlag
                    Case "F"    '## 정량
                        strIntBase = strIntBase & "N"
                        strResult = strIntResult
                    Case "I"    '## 정성
                        strIntBase = strIntBase & "C"
                        '-- Edit by Sewon,Oh(2007.08.02)
                        'Anti-HCV(전주예수병원 요구사항)
                        '장비판정값을 사용하지 않고 자체적으로 관리한다.
                        If strIntBase = "385C" Then
                            Select Case CDbl(strResult)
                                '2018.09.17 수정
                                Case Is < 1: strResult = "N"
                                Case 1 To 5: strResult = "WeaklyPositive"
                                Case Is > 5: strResult = "P"
                            End Select
                            strIntResult = strResult
                        '2018.11.15 추가 : HAVIgM
                        ElseIf strIntBase = "800C" Then
                            Select Case CDbl(strResult)
                                Case Is < 0.8: strResult = "N"
                                Case 0.8 To 1.21: strResult = "WeaklyPositive"
                                Case Is > 1.21: strResult = "P"
                            End Select
                            strIntResult = strResult
                        Else
                            Select Case Mid$(strIntResult, 1, 1)
                                '2018.09.17 수정
                                Case "N":   strResult = "N"
                                Case "G":   strResult = "Grayzone"
                                Case "R":   strResult = "P"
                                Case "P":   strResult = "P"
                                
                            End Select
                        End If
                    

                    '-- Edit by Sewon,Oh(2016.10.26)
                    'HbsAb2(Anti-Hbe)
                    '장비판정값이 없어서 정량값으로 판정한다.
                    '>=10 : Positive
                    '< 10 : Negative
                    Case "P"
                        'HbsAb2(Anti-Hbe)
                        If strIntBase = "694" Then
                            strIntBase = strIntBase & "C"
                            If IsNumeric(strResult) Then
                                Select Case CCur(strResult)
                                    Case Is < 10:  strResult = "Negative"
                                    Case Is >= 10: strResult = "Positive"
                                    '2018.09.17 수정
                                    Case Is < 10:  strResult = "N"
                                    Case Is >= 10: strResult = "P"
                                End Select
                                strIntResult = strResult
                            Else
                                'strResult = "Positive"
                                '2018.09.17 수정
                                strResult = "P"
                                strIntResult = strResult
                            End If
                        End If
                        
                        
                End Select
                
                '매독검사
                If strIntBase = "565C" Then
                    If UCase(strResult) = "NEGATIVE" Then
                        '2018.09.17 수정
                        strResult = "N"
                    ElseIf UCase(strResult) = "POSITIVE" Then
                        'strResult = "Reactive"
                        '2018.09.17 수정
                        strResult = "P"
                    End If
                Else
                    If UCase(strResult) = "NONREACTIVE" Then
                        'strResult = "Negative"
                        '2018.09.17 수정
                        strResult = "N"
                    ElseIf UCase(strResult) = "REACTIVE" Then
                        'strResult = "Positive"
                        '2018.09.17 수정
                        strResult = "P"
                    End If
                
                    If UCase(strIntResult) = "NONREACTIVE" Then
                        'strIntResult = "Negative"
                        '2018.09.17 수정
                        strIntResult = "N"
                    ElseIf UCase(strIntResult) = "REACTIVE" Then
                        'strIntResult = "Positive"
                        '2018.09.17 수정
                        strIntResult = "P"
                    End If
                End If
                
                If objIntNms.ExistIntBase(strIntBase) Then
                    Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                         strIntResult, strResult, strRackPos)
                    mIntLib.State = "R"
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
                            Call SendOrder
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
'-----------------------------------------------------------------------------'
'   기능 : 결과판정, 결과저장, 화면표시
'   인수 :
'       - pIntInfo : 인터페이스 검체정보 클래스
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pIntInfo As clsIISIntInfo)
    Dim objAccInfo  As clsIISAccInfo    '접수내역 클래스
    Dim vBarNo      As Variant 'Spread의 바코드번호
    Dim strBarNo    As String  '바코드번호
    Dim strSpcYY    As String  '검체연도
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
        strSpcYY = Mid$(strBarNo, 1, SPCYYLEN)
        lngSpcNo = CLng(Mid$(strBarNo, SPCYYLEN + 1, SPCNOLEN))
        Set objAccInfo = mIntLib.AccInfos(strSpcYY, lngSpcNo)

        Call SetComplete2(objAccInfo)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Set objAccInfo = Nothing

        '## ClientDb, Server에 결과저장
        Call mIntLib.SaveClientDb(strSpcYY, lngSpcNo)
        Call mIntLib.SaveResult(strSpcYY, lngSpcNo)
        '190022942188
        Call mIntLib.Remove(strSpcYY, lngSpcNo)
        
        StatusBar.Panels(2).Text = "검체번호:" & strBarNo & " 를 정상적으로 결과저장 했습니다."
    End If

    '## tblReady에서 전송된 검체삭제
    With tblReady
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
            If Mid(CStr(vBarNo), 1, SPCYYLEN + SPCNOLEN) = Mid(strBarNo, 1, SPCYYLEN + SPCNOLEN) Then
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
            'slSender.Add('1H|\^&||||||||||P|1|'+CR+ETX);
            
            '## 접수정보 유무를 판단하여 SndPhase변경
            If mOrder.NoOrder = True Then
                '## 접수정보가 없는경우
                mIntLib.SndPhase = 3
            Else
                mIntLib.SndPhase = 2
            End If

        Case 2  '## Patient
            'strOutput = mIntLib.GetFrameNo & "P|1||" & mOrder.AccInfo.PtId & "||" & vbCr & ETX
            strOutput = mIntLib.GetFrameNo & "P|1||" & mOrder.AccInfo.PtId & "|" & Han2EngOCX1.HanToEng(mOrder.AccInfo.Name) & "|" & vbCr & ETX
            'slSender.Add('2P|1|'+TMaster.FPID+'|||'+CR+ ETX);
            mIntLib.SndPhase = 4

        Case 3  '## No Order
            strOutput = mIntLib.GetFrameNo & "Q|1|^" & mOrder.BarNo & "||^^^ALL||||||||X" & vbCr & ETX
            mIntLib.SndPhase = 5

        Case 4  '## Order
            With mOrder
                If .IsSending = False Then  '## 최초 보낼때
                    strOutput = "O|1|" & .BarNo & "||" & .GetOrder & "|R||||||N||||||||||||||Q"
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
            End With

        Case 5  '## Termianator
            strOutput = mIntLib.GetFrameNo & "L|1" & vbCr & ETX
            mIntLib.SndPhase = 6

        Case 6  '## EOT
            mIntLib.State = ""
            wSck.SendData (EOT)
            Call mIntLib.WriteLog(EOT, ccPCLog)
            Exit Sub
    End Select

    strOutput = STX & strOutput & mOrder.GetChkSum(strOutput) & vbCrLf
    Debug.Print strOutput
    wSck.SendData (strOutput)
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
            .Text = objIntResult.IntNm & DIV & objIntResult.IntResult
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
        
        .Col = TCompleteEnum.ccNo:      .Text = strRackPos 'pAccInfo.SpcPos
        .Col = TCompleteEnum.ccBarNo:   .Text = pAccInfo.GetBarNo
        .Col = TCompleteEnum.ccAccNo:   .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)

        i = 0
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
                If objResult.RstCd <> "" Then
                    If .MaxCols <= .DataColCnt Then
                        .MaxCols = .MaxCols + 1
                    End If
                    .Col = TCompleteEnum.ccResult + i
                    .ColHidden = True
                    .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
                            DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & DIV & _
                            IIf(objResult.Ref.RefFg = "1", mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal), "")
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
            Error.SetLog App.EXEName, "frmIISAlinity", "GetEqpComm", strErrMsg, Now
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
        Error.SetLog App.EXEName, "frmIISAlinity", "GetEqpComm", strErrMsg, Now
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
    Dim strSpcYY    As String   '검체연도
    Dim lngSpcNo    As Long     '검체번호

    Select Case vMenuID
        Case DELETE     '## Delete
            With tblReady
                Call .GetText(TReadyEnum.ccBarNo, .ActiveRow, vBarNo)
                If vBarNo <> "" Then
                    strSpcYY = Mid$(vBarNo, 1, SPCYYLEN)
                    lngSpcNo = CLng(Mid$(vBarNo, SPCYYLEN + 1, SPCNOLEN))
                    Call mIntLib.AccInfos.Remove(strSpcYY, lngSpcNo)
                    Call .DeleteRows(.ActiveRow, 1)
                    Call mTblClear(tblResult)
                End If
            End With
        Case DELETEALL  '## Delete All
            Call mIntLib.AccInfos.RemoveAll
            Call mTblClear(tblReady)
    End Select
End Sub

Private Sub tmrNow_Timer()
    Dim AdoRS       As ADODB.Recordset
    Dim strSQL      As String
    Dim strBarNo    As String
    

'QC : LOW = 0327121801,
'       N = 0327121802,
'       H = 0327121803  12:Q
                '3O|1|0327121801|0327121801^K2472^1^1^5|^^^593^CA 125 II^UNDILUTED|R||||||Q||||||||||||||F
                
    If gNow = "" And gNow <> Format(Now, "yyyymmdd") Then
        gNow = Format(Now, "yyyymmdd")
        
        strBarNo = gLow '"0327121801"
        strSQL = ""
        strSQL = strSQL & ""
        strSQL = strSQL & "SELECT DISTINCT b.SPCYY, b.SPCNO " & vbCrLf
        strSQL = strSQL & "  FROM s2lab026 a, s2lab201 b    " & vbCrLf
        strSQL = strSQL & " WHERE a.lotno = '" & strBarNo & "'" & vbCrLf
        strSQL = strSQL & "   AND a.ACCDT = '" & gNow & "'  " & vbCrLf
        strSQL = strSQL & "   AND a.levelcd =  'L'          " & vbCrLf
        strSQL = strSQL & "   AND a.rstcd is null           " & vbCrLf
        strSQL = strSQL & "   AND a.WORKAREA = b.WORKAREA   " & vbCrLf
        strSQL = strSQL & "   AND a.ACCDT    = b.ACCDT      " & vbCrLf
        strSQL = strSQL & "   AND a.ACCSEQ   = b.ACCSEQ     " & vbCrLf
        
        Call mIntLib.WriteLog(strSQL, ccEqp)
        
        Set AdoRS = New ADODB.Recordset
        Set AdoRS = DbCon.Execute(strSQL, , adCmdText)
        If Not (AdoRS.BOF Or AdoRS.EOF) Then
            gSpcYY_L = AdoRS.Fields("SPCYY").Value & ""
            gSpcNo_L = AdoRS.Fields("SPCNO").Value & ""
            gBarno_L = gSpcYY_L & Format$(gSpcNo_L, String$(SPCNOLEN, "0"))
        End If
        Set AdoRS = Nothing
    
        strBarNo = gNormal '"0327121802"
        strSQL = ""
        strSQL = strSQL & ""
        strSQL = strSQL & "SELECT DISTINCT b.SPCYY, b.SPCNO " & vbCrLf
        strSQL = strSQL & "  FROM s2lab026 a, s2lab201 b    " & vbCrLf
        strSQL = strSQL & " WHERE a.lotno = '" & strBarNo & "'" & vbCrLf
        strSQL = strSQL & "   AND a.ACCDT = '" & gNow & "'" & vbCrLf
        strSQL = strSQL & "   AND a.levelcd =  'N'      " & vbCrLf
        strSQL = strSQL & "   AND a.rstcd is null           " & vbCrLf
        strSQL = strSQL & "   AND a.WORKAREA = b.WORKAREA   " & vbCrLf
        strSQL = strSQL & "   AND a.ACCDT    = b.ACCDT      " & vbCrLf
        strSQL = strSQL & "   AND a.ACCSEQ   = b.ACCSEQ     " & vbCrLf
        
        Call mIntLib.WriteLog(strSQL, ccEqp)
        
        Set AdoRS = New ADODB.Recordset
        Set AdoRS = DbCon.Execute(strSQL, , adCmdText)
        If Not (AdoRS.BOF Or AdoRS.EOF) Then
            gSpcYY_N = AdoRS.Fields("SPCYY").Value & ""
            gSpcNo_N = AdoRS.Fields("SPCNO").Value & ""
            gBarno_N = gSpcYY_N & Format$(gSpcNo_N, String$(SPCNOLEN, "0"))
        End If
        Set AdoRS = Nothing
    
        strBarNo = gHigh '"0327121803"
        strSQL = ""
        strSQL = strSQL & ""
        strSQL = strSQL & "SELECT DISTINCT b.SPCYY, b.SPCNO " & vbCrLf
        strSQL = strSQL & "  FROM s2lab026 a, s2lab201 b    " & vbCrLf
        strSQL = strSQL & " WHERE a.lotno = '" & strBarNo & "'" & vbCrLf
        strSQL = strSQL & "   AND a.ACCDT = '" & gNow & "'" & vbCrLf
        strSQL = strSQL & "   AND a.levelcd =  'H'      " & vbCrLf
        strSQL = strSQL & "   AND a.rstcd is null           " & vbCrLf
        strSQL = strSQL & "   AND a.WORKAREA = b.WORKAREA   " & vbCrLf
        strSQL = strSQL & "   AND a.ACCDT    = b.ACCDT      " & vbCrLf
        strSQL = strSQL & "   AND a.ACCSEQ   = b.ACCSEQ     " & vbCrLf
        
        Call mIntLib.WriteLog(strSQL, ccEqp)
        
        Set AdoRS = New ADODB.Recordset
        Set AdoRS = DbCon.Execute(strSQL, , adCmdText)
        If Not (AdoRS.BOF Or AdoRS.EOF) Then
            gSpcYY_H = AdoRS.Fields("SPCYY").Value & ""
            gSpcNo_H = AdoRS.Fields("SPCNO").Value & ""
            gBarno_H = gSpcYY_H & Format$(gSpcNo_H, String$(SPCNOLEN, "0"))
        End If
        Set AdoRS = Nothing
    End If
                
End Sub

Private Sub txtHigh_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If txtHigh.Text <> "" Then
            Call WritePrivateProfileString("ALINITY", "H", txtHigh.Text, App.Path & "\Alinity.ini")
        End If
    End If

End Sub

Private Sub txtLow_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If txtLow.Text <> "" Then
            Call WritePrivateProfileString("ALINITY", "L", txtLow.Text, App.Path & "\Alinity.ini")
        End If
    End If
    
End Sub

Private Sub txtNormal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If txtNormal.Text <> "" Then
            Call WritePrivateProfileString("ALINITY", "N", txtNormal.Text, App.Path & "\Alinity.ini")
        End If
    End If

End Sub

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



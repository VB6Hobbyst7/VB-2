VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Begin VB.Form frmIISUF1000i 
   BackColor       =   &H00DBE6E6&
   Caption         =   "UF1000i"
   ClientHeight    =   9180
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   3960
      TabIndex        =   26
      Top             =   135
      Visible         =   0   'False
      Width           =   1410
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
      SpreadDesigner  =   "frmIISUF1000i.frx":0000
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
      SpreadDesigner  =   "frmIISUF1000i.frx":04FA
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
      SpreadDesigner  =   "frmIISUF1000i.frx":0D3C
      TextTip         =   2
   End
End
Attribute VB_Name = "frmIISUF1000i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISCoa1.frm
'   작성자  : 오세원
'   내  용  : Stago 장비폼
'   작성일  : 2008-07-08
'   버  전  :
'   병  원  :
'       1. 안동성소병원
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

Private strBuf As String
Private strBarNum As String
Private strPosNum As String

Private s0201 As String
Private s0202 As String
Private s0100 As String
Private s0000 As String
Private s0401 As String

Private s0300 As String
Private s0402 As String
Private s00D9 As String
Private s0107 As String
Private s00DA As String
Private s0501 As String
Private s0502 As String

Private lsFlagRBC       As String
Private lsFlagWBC       As String
Private lsFlagEC        As String
Private lsFlagCAST      As String
Private lsFlagBACT      As String


Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property

Private Sub Command1_Click()


'DS4401050.00            UF-1000i^05366719^15358U20131212091626     2 2   130010901285I0100B****************07.8001.001972509442023982009418*            8    E                           
'DP4402050.00            UF-1000i^05366719^1535805020100352.80020200257.30010000053.00000000011.46040100054.30
'DC4403050.00            UF-1000i^05366719^153580300D901070402
'DQ4404050.00            UF-1000i^05366719^153580700D900008.06010700039.60050100000.00030000000.10040200289.1000DA00001.98050200010.80
'DD4405050.00            UF-1000i^05366719^15358030C00000000000C01000000020C0200000001


    strBuf = "DS4401050.00            UF-1000i^05366719^15358U20131212091626     2 2   130010901285I0100B****************07.8001.001972509442023982009418*            8    E                           "
    strBuf = "DP4402050.00            UF-1000i^05366719^1535805020100352.80020200257.30010000053.00000000011.46040100054.30"
    strBuf = "DC4403050.00            UF-1000i^05366719^153580300D901070402"
    strBuf = "DQ4404050.00            UF-1000i^05366719^153580700D900008.06010700039.60050100000.00030000000.10040200289.1000DA00001.98050200010.80"
    strBuf = "DD4405050.00            UF-1000i^05366719^15358030C00000000000C01000000020C0200000001"

    
    strBuf = "DS4401050.00            UF-1000i^05366719^15358U20140123101536     7 2    14000106780I0100B****************07.8001.000212865479002491129441    *          B                              "
    strBuf = "DP4402050.00            UF-1000i^05366719^1535805020100197.30020200000.00010000006.70000000000.99040100003.90"
    strBuf = "DC4403050.00            UF-1000i^05366719^1535800"
    strBuf = "DQ4404050.00            UF-1000i^05366719^153580700D900000.99010700000.00050100000.00030000000.00040200000.0000DA00000.00050200021.40"
    strBuf = "DD4405050.00            UF-1000i^05366719^15358030C00000000010C01000000030C0200000000"

    strBuf = "DS4401050.00            UF-1000i^05366719^15358U20140123101611     7 3    14000107888I0000B****************07.8001.001238211132015104011089                                              "
    strBuf = "DP4402050.00            UF-1000i^05366719^1535805020100386.60020200133.70010000007.30000000003.67040100333.90"
    strBuf = "DC4403050.00            UF-1000i^05366719^1535800"
    strBuf = "DQ4404050.00            UF-"
    
    strBuf = "DS4401050.00            UF-1000i^05366719^15358U20140124091055     2 2    14000110265I0100B****************07.8001.003910131876046578031493*            8   DE                           "
    strBuf = "DP4402050.00            UF-1000i^05366719^1535805020100048.60020200406.00010000009.40000000006.79040100420.80"
    strBuf = "DC4403050.00            UF-1000i^05366719^153580300D901070402"
    strBuf = "DQ4404050.00            UF-1000i^05366719^153580700D900008.06010700039.60050100000.00030000000.10040200289.1000DA00001.98050200010.80"
    strBuf = "DD4405050.00            UF-1000i^05366719^15358030C00000000000C01000000020C0200000001"
    
    Call MSComm_OnComm
    



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
    Call GetEqpComm
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
    Set frmIISUF1000i = Nothing
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

            '## 결과표시 버그수정
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
'GoTo RST
    Select Case MSComm.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long

            Buffer = MSComm.Input
'RST:
'            Buffer = strBuf
            
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
                            Case ETX
                                mIntLib.Phase = 1
                                MSComm.Output = ACK
                                Call EditRcvData
                            Case Else
                                Call mIntLib.AddBuffer(BufChar)
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
    Dim objIntNm    As clsIISIntNm      '장비별 검사항목 클래스

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
    Dim intTestCnt   As Integer
    Dim i As Integer
    Dim j As Integer
    
    Set objIntNms = mIntLib.IntNms
    mIntLib.State = "R"
    
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        strType = Mid$(strRcvBuf, 1, 2)

        Select Case strType
            Case "DS"    '## Order
                lsFlagRBC = ""
                lsFlagWBC = ""
                lsFlagEC = ""
                lsFlagCAST = ""
                lsFlagBACT = ""
                
                strBarNo = Mid(strRcvBuf, 74, 12)
                'osw add
                strBarNo = Format(strBarNo, "#")
                strSeg = Trim(Mid(strRcvBuf, 67, 2))
                strPos = Trim(Mid(strRcvBuf, 69, 4))
                
                'DS4401050.00            UF-1000i^05366719^15358U20140124091055     2 2    14000110265I0100B********
                '********07.8001.003910131876046578031493*            8   DE                           
                'Set objIntInfo = New clsIISIntInfo
                'With objIntInfo
                '    .BarNo = strBarNo
                '    .SpcPos = strPos & "/" & strSeg
                'End With
                strBarNum = strBarNo
                strPosNum = strPos & "/" & strSeg
                'Call GetOrder(strBarNo)
                
                lsFlagRBC = Mid(strRcvBuf, 140, 1)
                Select Case lsFlagRBC
                    Case "*":  lsFlagRBC = "(*)" '"Low reliability"
                    Case "+":  lsFlagRBC = "(+)" '"Positivie"
                    Case Else: lsFlagRBC = ""
                End Select
                
                lsFlagWBC = Mid(strRcvBuf, 141, 1)
                Select Case lsFlagWBC
                    Case "*":  lsFlagWBC = "(*)" '"Low reliability"
                    Case "+":  lsFlagWBC = "(+)" '"Positivie"
                    Case Else: lsFlagWBC = ""
                End Select
                
                lsFlagEC = Mid(strRcvBuf, 142, 1)
                Select Case lsFlagEC
                    Case "*":  lsFlagEC = "(*)" '"Low reliability"
                    Case "+":  lsFlagEC = "(+)" '"Positivie"
                    Case Else: lsFlagEC = ""
                End Select
            
                lsFlagCAST = Mid(strRcvBuf, 143, 1)
                Select Case lsFlagCAST
                    Case "*":  lsFlagCAST = "(*)" '"Low reliability"
                    Case "+":  lsFlagCAST = "(+)" '"Positivie"
                    Case Else: lsFlagCAST = ""
                End Select
                
                lsFlagBACT = Mid(strRcvBuf, 144, 1)
                Select Case lsFlagBACT
                    Case "*":  lsFlagBACT = "(*)" '"Low reliability"
                    Case "+":  lsFlagBACT = "(+)" '"Positivie"
                    Case Else: lsFlagBACT = ""
                End Select
                
            Case "DP"    '## Result
                strTemp = Mid(strRcvBuf, 50)
            
                For intTestCnt = 1 To 5
                    strIntBase = Mid(strTemp, 1, 4)

                    Select Case strIntBase
                        Case "0201" 'RBC
                            'strIntBase = Mid(strTemp, 1, 4)
                            strIntResult = Mid(strTemp, 5, 8)
                            strIntResult = Format(CCur(strIntResult), "#0.00")
                            strResult = strIntResult
                            strIntResult = Format(CCur(strIntResult) * 0.18, "#0.00")
                            
                            If CCur(strIntResult) <= 1 Then
                                 strResult = "1개미만"
                            ElseIf CCur(strIntResult) <= 5 Then
                                 strResult = "1-4개"
                            ElseIf CCur(strIntResult) <= 10 Then
                                 strResult = "5-9개"
                            ElseIf CCur(strIntResult) <= 20 Then
                                 strResult = "10-19개"
                            ElseIf CCur(strIntResult) <= 30 Then
                                 strResult = "20-29개"
                            ElseIf CCur(strIntResult) <= 50 Then
                                 strResult = "30-49개"
                            ElseIf CCur(strIntResult) <= 100 Then
                                 strResult = "50-99개"
                            Else
                                 strResult = "100개이상"
                            End If
                            
'                            If lsFlagRBC = "" Then
'                                s0201 = strResult
'                            Else
'                                s0201 = lsFlagRBC
'                            End If
                            
                            s0201 = lsFlagRBC & " " & strResult
                            
                        Case "0202" 'WBC
                            'strIntBase = Mid(strTemp, 1, 4)
                            strIntResult = Mid(strTemp, 5, 8)
                            strIntResult = Format(CCur(strIntResult), "#0.00")
                            strResult = strIntResult
                            strIntResult = Format(CCur(strIntResult) * 0.18, "#0.00")

                            If CCur(strIntResult) <= 1 Then
                                 strResult = "1개미만"
                            ElseIf CCur(strIntResult) <= 5 Then
                                 strResult = "1-4개"
                            ElseIf CCur(strIntResult) <= 10 Then
                                 strResult = "5-9개"
                            ElseIf CCur(strIntResult) <= 20 Then
                                 strResult = "10-19개"
                            ElseIf CCur(strIntResult) <= 30 Then
                                 strResult = "20-29개"
                            ElseIf CCur(strIntResult) <= 50 Then
                                 strResult = "30-49개"
                            ElseIf CCur(strIntResult) <= 100 Then
                                 strResult = "50-99개"
                            Else
                                 strResult = "100개이상"
                            End If
                            
'                            If lsFlagWBC = "" Then
'                                s0202 = strResult
'                            Else
'                                s0202 = lsFlagRBC
'                            End If
                            
                            s0202 = lsFlagWBC & " " & strResult
                        
                        Case "0100" 'EC
                            strIntResult = Mid(strTemp, 5, 8)
                            strIntResult = Format(CCur(strIntResult), "#0.00")
                            strResult = strIntResult
                            strIntResult = Format(CCur(strIntResult) * 0.18, "#0.00")
                            
                            '-- 2014.08.04 수정
'''                            If CCur(strIntResult) <= 1 Then
'''                                 strResult = ""
'''                            ElseIf CCur(strIntResult) <= 5 Then
'''                                 strResult = "FEW"
'''                            ElseIf CCur(strIntResult) <= 20 Then
'''                                 strResult = "SOME"
'''                            ElseIf CCur(strIntResult) <= 100 Then
'''                                 strResult = "MANY"
'''                            Else
'''                                 strResult = "VERY MANY"
'''                            End If
                            
                            If CCur(strIntResult) <= 1 Then
                                 strResult = "1개미만"
                            ElseIf CCur(strIntResult) <= 5 Then
                                 strResult = "1-4개"
                            ElseIf CCur(strIntResult) <= 10 Then
                                 strResult = "5-9개"
                            ElseIf CCur(strIntResult) <= 20 Then
                                 strResult = "10-19개"
                            ElseIf CCur(strIntResult) <= 30 Then
                                 strResult = "20-29개"
                            ElseIf CCur(strIntResult) <= 50 Then
                                 strResult = "30-49개"
                            ElseIf CCur(strIntResult) <= 100 Then
                                 strResult = "50-99개"
                            Else
                                 strResult = "100개이상"
                            End If
                            
'                            If lsFlagWBC = "" Then
'                                s0100 = strResult
'                            Else
'                                s0100 = lsFlagEC
'                            End If
                        
                            s0100 = lsFlagEC & " " & strResult
                        
'                        Case "0000" 'CAST
'                            strIntResult = Mid(strTemp, 5, 8)
'                            strIntResult = Format(CCur(strIntResult), "#0.0")
'                            strResult = strIntResult
'                            strIntResult = Format(CCur(strIntResult) * 0.18, "#0.0")
'
'                            If CCur(strIntResult) <= 1 Then
'                                 strResult = "1개미만"
'                            ElseIf CCur(strIntResult) < 5 Then
'                                 strResult = "1-4"
'                            ElseIf CCur(strIntResult) < 10 Then
'                                 strResult = "5-9"
'                            ElseIf CCur(strIntResult) < 20 Then
'                                 strResult = "10-19"
'                            Else
'                                 strResult = "20개 이상"
'                            End If
'
'                            s0000 = strResult
                            
                        Case "0401" 'BACT
                            strIntResult = Mid(strTemp, 5, 8)
                            strIntResult = Format(CCur(strIntResult), "#0.00")
                            strResult = strIntResult
                            strIntResult = Format(CCur(strIntResult) * 0.18, "#0.00")
                            
                            If CCur(strIntResult) < 0.9 Then '5 ul
                                strResult = ""
                            ElseIf CCur(strIntResult) <= 18 Then
                                 strResult = "FEW"
                            ElseIf CCur(strIntResult) <= 180 Then
                                 strResult = "SOME"
                            ElseIf CCur(strIntResult) <= 900 Then
                                 strResult = "MANY"
                            'ElseIf CCur(strIntResult) < 100 Then
                            '     strResult = "MANY"
                            Else
                                 strResult = "VERY MANY"
                            End If
                            
'                            If lsFlagWBC = "" Then
'                                s0401 = strResult
'                            Else
'                                s0401 = lsFlagBACT
'                            End If
                    
                            s0401 = lsFlagBACT & " " & strResult
                    
                    End Select
                    
                    'Call objIntInfo.IntResults.Add(strIntBase, objIntNm.IntNm, strIntResult, strResult)
                    strResult = ""
                    strTemp = Mid(strTemp, 13)
                Next
                
            Case "DC"    '## Result
                                                '-- 끝에 두자리 : 양성갯수
'DC4403050.00            UF-1000i^05366719^1535803
'-- 양성 검사채널
'00D901070402
                
                Set objIntInfo = New clsIISIntInfo
                With objIntInfo
                    .BarNo = strBarNum
                    .SpcPos = strPosNum
                End With
                Call GetOrder(strBarNum)
                
                intTestCnt = Mid(strRcvBuf, 48, 2)
                strTemp = Mid(strRcvBuf, 50)
                
                For Each objIntNm In mIntLib.IntNms
                    strIntBase = objIntNm.IntBase
                    Select Case strIntBase
                        Case "0201"
                            strResult = s0201
                        Case "0202"
                            strResult = s0202
                        Case "0100"
                            strResult = s0100
                        Case "0000"
                            strResult = s0000
                        Case "0401"
                            strResult = s0401
                        

                        Case "0300"
                            strResult = ""
                            If intTestCnt > 0 Then
                                For i = 1 To intTestCnt
                                    j = i - 1
                                    If strIntBase = Mid(strTemp, (1 + (j * 4)), 4) Then
                                        strResult = "+"
                                        Exit For
                                    End If
                                Next
                            End If
                        Case "0402"
                            strResult = ""
                            If intTestCnt > 0 Then
                                For i = 1 To intTestCnt
                                    j = i - 1
                                    If strIntBase = Mid(strTemp, (1 + (j * 4)), 4) Then
                                        strResult = "+"
                                        Exit For
                                    End If
                                Next
                            End If
                        Case "00D9"
                            strResult = ""
                            If intTestCnt > 0 Then
                                For i = 1 To intTestCnt
                                    j = i - 1
                                    If strIntBase = Mid(strTemp, (1 + (j * 4)), 4) Then
                                        strResult = "+"
                                        Exit For
                                    End If
                                Next
                            End If
                        Case "0107"
                            strResult = ""
                            If intTestCnt > 0 Then
                                For i = 1 To intTestCnt
                                    j = i - 1
                                    If strIntBase = Mid(strTemp, (1 + (j * 4)), 4) Then
                                        strResult = "+"
                                        Exit For
                                    End If
                                Next
                            End If
                        Case "00DA"
                            strResult = ""
                            If intTestCnt > 0 Then
                                For i = 1 To intTestCnt
                                    j = i - 1
                                    If strIntBase = Mid(strTemp, (1 + (j * 4)), 4) Then
                                        strResult = "+"
                                        Exit For
                                    End If
                                Next
                            End If
                        Case "0501"
                            strResult = ""
                            If intTestCnt > 0 Then
                                For i = 1 To intTestCnt
                                    j = i - 1
                                    If strIntBase = Mid(strTemp, (1 + (j * 4)), 4) Then
                                        strResult = "+"
                                        Exit For
                                    End If
                                Next
                            End If
                        Case "0502"
                            strResult = ""
                            If intTestCnt > 0 Then
                                For i = 1 To intTestCnt
                                    j = i - 1
                                    If strIntBase = Mid(strTemp, (1 + (j * 4)), 4) Then
                                        strResult = "+"
                                        Exit For
                                    End If
                                Next
                            End If

                    End Select
                    
                    Call objIntInfo.IntResults.Add(strIntBase, objIntNm.IntNm, strIntResult, strResult)
                    strResult = ""
                    mIntLib.State = "R"
                Next

                lsFlagRBC = ""
                lsFlagWBC = ""
                lsFlagEC = ""
                lsFlagCAST = ""
                lsFlagBACT = ""

                If mIntLib.State = "R" Then
                    Call SaveServer(objIntInfo)
                    Set objIntInfo = Nothing
                    mIntLib.State = ""
                End If

            
'            Case "L"    '## Terminator
'                '## DB에 결과저장
'                If mIntLib.State = "R" Then
'                    Call SaveServer(objIntInfo)
'                    Set objIntInfo = Nothing
'                    mIntLib.State = ""
'                End If
        End Select
    Next
    
    Set objIntNm = Nothing
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

'MsgBox mIntLib.SndPhase
    Select Case mIntLib.SndPhase
        Case 1  '## Header
'            strOutput = mIntLib.GetFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            strOutput = mIntLib.GetFrameNo & "H|\^&|||99^2.00" & vbCr & ETX
            
            '## 접수정보 유무를 판단하여 SndPhase변경
            If mOrder.NoOrder = True Then
                '## 접수정보가 없는경우
                mIntLib.SndPhase = 3
            Else
                mIntLib.SndPhase = 2
            End If

        Case 2  '## Patient
            'strOutput = mIntLib.GetFrameNo & "P|1||" & mOrder.AccInfo.PtId & "||" & vbCr & ETX
            strOutput = mIntLib.GetFrameNo & "P|1|||" & mOrder.AccInfo.PtId & "|^1^1^56|||19700505" & vbCr & ETX
            mIntLib.SndPhase = 4

        Case 3  '## No Order
            strOutput = mIntLib.GetFrameNo & "Q|1|^" & mOrder.BarNo & "||^^^ALL||||||||X" & vbCr & ETX
            mIntLib.SndPhase = 5

        Case 4  '## Order
            With mOrder
                If .IsSending = False Then  '## 최초 보낼때
'                    strOutput = "O|1|" & .BarNo & "||" & .GetOrder & "|R||||||N||||||||||||||Q"
                    strOutput = "O|1|" & .BarNo & "||" & .GetOrder & "|R" '||||||N||||||||||||||Q"
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
'            strOutput = mIntLib.GetFrameNo & "L|1" & vbCr & ETX
            strOutput = mIntLib.GetFrameNo & "L|1N" & vbCr & ETX
            mIntLib.SndPhase = 6

        Case 6  '## EOT
            mIntLib.State = ""
            MSComm.Output = Chr(4)   'EOT
            Call mIntLib.WriteLog(EOT, ccPCLog)
            Exit Sub
    End Select

    strOutput = STX & strOutput & mOrder.GetChkSum(strOutput) & vbCrLf
    MSComm.Output = strOutput
    Debug.Print strOutput
    
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
        
        .Col = TCompleteEnum.ccNo:      .Text = pAccInfo.SpcPos
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
            Error.SetLog App.EXEName, "frmIISCoa1", "GetEqpComm", strErrMsg, Now
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
        Error.SetLog App.EXEName, "frmIISCoa1", "GetEqpComm", strErrMsg, Now
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




VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Begin VB.Form frmIISChorusTRIO 
   BackColor       =   &H00DBE6E6&
   Caption         =   "ChorusTRIO"
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
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   6548
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   109
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
      TabIndex        =   8
      Top             =   409
      Width           =   8595
      Begin MedControls1.LisLabel lblPtId 
         Height          =   315
         Left            =   1245
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
      TabIndex        =   4
      Top             =   8569
      Width           =   1215
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
      Left            =   1433
      TabIndex        =   1
      Text            =   "123456789011"
      Top             =   514
      Visible         =   0   'False
      Width           =   1530
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
      Left            =   5183
      MaxLength       =   9
      TabIndex        =   0
      Top             =   514
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      Height          =   495
      Left            =   11483
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8569
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   12698
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8569
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   6578
      Top             =   8434
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
   End
   Begin MDIControls.MDIActiveX MDIActiveX 
      Left            =   7238
      Top             =   8509
      _ExtentX        =   847
      _ExtentY        =   794
   End
   Begin FPSpread.vaSpread tblReady 
      Height          =   3570
      Left            =   105
      TabIndex        =   6
      Top             =   585
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   6297
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
      SpreadDesigner  =   "frmIISChorusTRIO.frx":0000
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4410
      Left            =   105
      TabIndex        =   7
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
      SpreadDesigner  =   "frmIISChorusTRIO.frx":04F2
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   98
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   109
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4249
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
      TabIndex        =   27
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
      MaxCols         =   7
      MaxRows         =   22
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISChorusTRIO.frx":0D42
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
      Left            =   203
      TabIndex        =   29
      Top             =   589
      Visible         =   0   'False
      Width           =   1200
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
      Left            =   3653
      TabIndex        =   28
      Top             =   589
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "frmIISChorusTRIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISiQ200.frm
'   작성자  : 오세원
'   내  용  : iQ200 장비폼
'   작성일  : 2007-10-26
'   버  전  :
'   병  원  :
'       1. 천안 단국대병원
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

Private mEqpCd  As String   '장비코드
Private mEqpKey As String   '장비키
Private lngNDXCnt  As Long
Private blnNDXCnt  As Boolean
Public strSpcNo As Integer

Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property

Private Sub Form_Activate()
    MainFrm.lblMenuNm = Me.Caption
    Me.MDIActiveX.WindowState = ccMaximize
End Sub

Private Sub Form_Load()
    Me.Caption = mEqpKey
    
    Me.MousePointer = vbHourglass
    
    Set mIntErrors = New clsIISIntErrors
    Set mIntLib = New clsIISInterface
    
    Call CtlClear
    Call mIntLib.SetConfig(mEqpCd, mEqpKey)
    Call GetEqpComm
    DoEvents
    lngNDXCnt = 0
    blnNDXCnt = False
    strSpcNo = 0
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Deactivate()
    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISChorusTRIO = Nothing
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
    Unload Me
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
            
            '## 1.0.4: 이상대(2005-07-12)
            '   - 화면표시 버그수정
            If tblResult.MaxRows <= tblResult.DataRowCnt Then
                tblResult.MaxRows = tblResult.MaxRows + 1
                tblResult.Row = tblResult.MaxRows
            Else
                tblResult.Row = tblResult.DataRowCnt + 1
            End If
            
            tblResult.Col = TResultEnum.ccTestNm:       tblResult.Text = mGetP(strTemp, TResultEnum.ccTestNm, DIV)
            tblResult.Col = TResultEnum.ccEqpResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccEqpResult, DIV)
            tblResult.Col = TResultEnum.ccLISResult:    tblResult.Text = mGetP(strTemp, TResultEnum.ccLISResult, DIV)
            tblResult.Col = TResultEnum.ccUnit:         tblResult.Text = mGetP(strTemp, TResultEnum.ccUnit, DIV)
            tblResult.Col = TResultEnum.ccHLDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccHLDiv, DIV)
            tblResult.Col = TResultEnum.ccDPDiv:        tblResult.Text = mGetP(strTemp, TResultEnum.ccDPDiv, DIV)
            tblResult.Col = TResultEnum.ccRef:          tblResult.Text = mGetP(strTemp, TResultEnum.ccRef, DIV)
        Next i
    End With
End Sub

Private Sub MSComm_OnComm()
    Dim EVMsg As String
    Dim ERMsg As String
    Dim Ret   As Long
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
    Dim strNDX      As String
    
    
    
    Select Case MSComm.CommEvent
        Case comEvReceive
            
            Buffer = MSComm.Input
            
            'Debug.Print Buffer
            Call mIntLib.WriteLog(Buffer, ccEqp)
            lngBufLen = Len(Buffer)
            
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                Select Case mIntLib.Phase
                    Case 1
                        Select Case BufChar
                            Case STX  'ENQ()
                                'MSComm.Output = Chr(6) 'ACK()
                                mIntLib.Phase = 2
                                mIntLib.BufCnt = 1
                        End Select
                    Case 2
                        Select Case BufChar
                            Case STX
                                mIntLib.ClearBuffer
                            Case ENQ
                                MSComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5) 'stx soh eot eeq
                                Debug.Print ">>" & Chr(2) & Chr(1) & Chr(4) & Chr(5)
                                mIntLib.ClearBuffer
                                mIntLib.Phase = 3
                            
                            'Case "D"
                            'Case Asc(vbCr)
                            '    MSComm.Output = Chr(6)
                            '    mIntLib.BufCnt = mIntLib.BufCnt + 1
                            '    mIntLib.Phase = 3
                            'Case Else
                            '    Call mIntLib.AddBuffer(BufChar)
                        End Select
                    Case 3
                        Select Case BufChar
'                            Case "N"
'                                strNDX = "N"
'                            Case "D"
'                                strNDX = strNDX & "D"
'                            Case "X"
'                                strNDX = strNDX & "X"
'                                If strNDX = "NDX" Then
'                                    lngNDXCnt = lngNDXCnt + 1
'                                    MSComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5) 'stx soh eot eeq
'        '                            Call EditRcvData
'                                    mIntLib.ClearBuffer
'                                End If
'                            Case "E5"
'                                'If lngNDXCnt > 0 Then
'                                    MSComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)  'stx soh eot eeq
'        '                            Call EditRcvData
'                                    mIntLib.ClearBuffer
'                                    mIntLib.Phase = 1
'                                    lngNDXCnt = 0
                                'End If
                                'mIntLib.Phase = 4
                            
'                            Case ""
'                                If blnNDXCnt = True Then
'                                    'MSComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)  'stx soh eot eeq
'                                    MSComm.Output = Chr(2) & "1" & Chr(4) & "37"
'                                    blnNDXCnt = False
'                                End If
                            Case STX
                                mIntLib.ClearBuffer
                                'mIntLib.Phase = 1
                                lngNDXCnt = 0
                                blnNDXCnt = True
                                
                            Case Else
                                Call mIntLib.AddBuffer(BufChar)
                        End Select
                        
                        lngNDXCnt = lngNDXCnt + 1
                        If lngNDXCnt >= 51 Then
                            MSComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5)  'stx soh eot eeq
                            Call EditRcvData
                            mIntLib.ClearBuffer
                            'mIntLib.Phase = 1
                            lngNDXCnt = 0
                            blnNDXCnt = True
                        
                        End If

                End Select
            Next
            
            MSComm.Output = Chr(2) & Chr(1) & Chr(4) & Chr(5) 'stx soh eot eeq
            
         '   MSComm.Output = Chr(2) & "1" & Chr(4) & "37"
'Timer1.Interval = 2000
'            Timer1.Enabled = True
            
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
    Dim objBuffer    As clsIISBuffer     '버퍼 클래스
    
    Dim vWorkNo      As Variant  'Spread의 WorkNo
    Dim vBarNo       As Variant  'Spread의 바코드번호
    Dim strRcvBuf    As String   '수신한 Data
    Dim strType      As String   '수신한 Record Type
    Dim strBarNo     As String   '수신한 BarNO
    Dim strWorkNo    As String   '수신한 WorkNo
    Dim strIntResult As String   '수신한 검사결과
    Dim strIntBase   As String   '수신한 장비기준 검사명
    Dim strResult    As String   'LIS 검사결과
    Dim strTemp      As String
    Dim i            As Long
    Dim varBarNo As Variant
    Dim intRow       As Integer
    Dim blnSameBar   As Boolean
    
    
    Set objIntNms = mIntLib.IntNms
    
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        
        Debug.Print strRcvBuf
        '위와 같은 포맷의 데이터가 발생하지 않으므로 필요없어졌음 - 오세원
        strType = Mid$(strRcvBuf, 3, 1)

'       ?2?20000307845       MYCO-M N0.1         Ndx

        blnSameBar = False
        
        If Mid(strRcvBuf, 1, 1) = "2" Then
            strBarNo = "1" & Mid(strRcvBuf, 3, 11)
            strBarNo = Format$(strBarNo, String$(SPCLEN, "#"))
            blnSameBar = True
        Else
            strBarNo = "1" & Mid(strRcvBuf, 4, 11)
            strBarNo = Format$(strBarNo, String$(SPCLEN, "#"))
            blnSameBar = True
        End If
        
        Set objIntInfo = New clsIISIntInfo
        
        strSpcNo = strSpcNo + 1
        
        With objIntInfo
            .BarNo = strBarNo
            .SpcPos = strSpcNo
        End With
        
'        With tblReady
'            .Col = 2
'            For intRow = 1 To .MaxRows
'                .Row = intRow
'                If Trim(.Text) = strBarNo Then
'                    blnSameBar = True
'                    Exit For
'                End If
'            Next
'            If blnSameBar = True Then
'                txtBarNo.Text = strBarNo
'                Call txtBarNo_KeyDown(vbKeyReturn, 0)
'            End If
'        End With
        
        If Mid(strRcvBuf, 1, 1) = "2" Then
            strIntBase = Mid(strRcvBuf, 21, 6)
            strIntResult = Val(Mid(strRcvBuf, 29, 6))
        Else
            strIntBase = Mid(strRcvBuf, 22, 6)
            strIntResult = Val(Mid(strRcvBuf, 30, 6))
        End If
        
        If strIntResult > 1.1 Then
            strIntResult = "POSITIVE:" & strIntResult
        ElseIf strIntResult < 0.9 Then
            strIntResult = "NEGATIVE:" & strIntResult
        Else
            strIntResult = "EQUIVOCAL:" & strIntResult
        End If

        '## 결과값에 "?", ">", "<" 포함되어 있으면 에러로 표시
'        If IsNumeric(strIntResult) Then
            strResult = strIntResult
'        Else
'            strResult = IISERROR
'        End If
        
        If objIntNms.ExistIntBase(strIntBase) Then
            Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                 strIntResult, strResult, "1")
        End If

        Call SaveServer(objIntInfo)
        Set objIntInfo = Nothing
    Next
    
'    mIntLib.Phase = 1
    Set objIntNms = Nothing
    Set objBuffer = Nothing
    Set objIntInfo = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : Unicode를 String값으로 변환
'   인수 : Unicode Value
'-----------------------------------------------------------------------------'
Private Function UniToStr(strUniValue As String) As String
    Dim intCnt          As Integer
    Dim strFr2UniVal    As String
    Dim strTo2UniVal    As String
    Dim strRsltUniVal As String

    For intCnt = 1 To Len(strUniValue) Step 4
        If Mid(strUniValue, intCnt, 4) = "0D0A" Then 'CarrageReturn + LineFeed
            UniToStr = UniToStr + vbCrLf
            strRsltUniVal = strRsltUniVal + "0000"
        Else
            strFr2UniVal = Mid(strUniValue, intCnt, 2)
            strTo2UniVal = Mid(strUniValue, intCnt + 2, 2)
            If strFr2UniVal = "00" Then 'English & Optional Character
                strRsltUniVal = strRsltUniVal + Mid(strUniValue, intCnt, 4)
                Select Case LCase(ChrW(Val("&H" & Mid(strRsltUniVal, intCnt, 4))))
                    '>> Skip
                    Case "a" To "z", 0 To 9, _
                         "-", "_", "=", "+", "\", "|", "[", "]", "{", "}", "<", ">", ",", ".", "/", "?", _
                         "~", "!", "@", "$", "#", "%", "^", "&", "*", "(", ")", ":", ";", "'", """", " "
                    Case Else
                        strRsltUniVal = Left(strRsltUniVal, Len(strRsltUniVal) - 4) + strTo2UniVal + strFr2UniVal
                End Select
            Else
                'Change Sequence
                strRsltUniVal = strRsltUniVal + strTo2UniVal + strFr2UniVal
            End If
            
            UniToStr = UniToStr + ChrW(Val("&H" & Mid(strRsltUniVal, intCnt, 4)))
        
        End If
    Next intCnt

End Function

''유니코드 -> 문자열
'Private Function UniToStr(Text As String) As String
'Dim a As String, b As String, c As String
'Dim i As Integer
'
'For i = 1 To Len(Text) Step 4
'    '리턴값 처리
'    If Mid(Text, i, 4) = "0D0A" Then
'        UniToStr = UniToStr + vbCrLf
'        c = c + "0000"
'    Else
'        '순서변경
'        a = Mid(Text, i, 2)
'        b = Mid(Text, i + 2, 2)
'        '영문/기타 문자일 경우, 다시 변경
'        If a = "00" Then
'            c = c + Mid(Text, i, 4)
'            '특수경우 제외
'            Select Case LCase(ChrW(Val("&H" & Mid(c, i, 4))))
'                Case "a" To "z", 0 To 9, "-", "_", "=", "+", "\", "|", "[", "]", "{", "}", "<", ">", ",", ".", "/", "?", "~", "!", "@", "$", "#", "%", "^", "&", "*", "(", ")", ":", ";", "'", """", " "
'                Case Else: c = Left(c, Len(c) - 4) + b + a
'            End Select
'        Else
'        c = c + b + a
'        End If
'        UniToStr = UniToStr + ChrW(Val("&H" & Mid(c, i, 4)))
'        'Debug.Print ChrW(Val("&H" & Mid(c, i, 4)))
'        'If ChrW(Val("&H" & Mid(c, i, 4))) = "&" Then Stop
'    End If
'Next i
'
'End Function

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
    
    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    If Not (objAccInfo Is Nothing) Then
        '## tblReady, tblresult, Label에 정보표시
        Call SetReady(objAccInfo)
        Call SetLabel(objAccInfo)
        Call SetResult(objAccInfo)
        
        Set objAccInfo = Nothing
    End If
    txtBarNo.Text = "": txtBarNo.SetFocus
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
            .Text = objIntResult.IntNm & DIV & objIntResult.Result
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
                        IIf(objResult.Ref.RefFg = "1", mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal), "")
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
            Error.SetLog App.EXEName, "frmIISUriscanPro", "GetEqpComm", strErrMsg, Now
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
        Error.SetLog App.EXEName, "frmIISUriscanPro", "GetEqpComm", strErrMsg, Now
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


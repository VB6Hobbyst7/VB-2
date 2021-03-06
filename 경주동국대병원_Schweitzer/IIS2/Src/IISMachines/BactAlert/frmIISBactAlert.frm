VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmIISBactAlert 
   BackColor       =   &H00DBE6E6&
   Caption         =   "BactAlert"
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
   Begin VB.CheckBox chkLoad 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Refresh"
      Height          =   285
      Left            =   8610
      TabIndex        =   27
      Top             =   8640
      Width           =   945
   End
   Begin VB.Timer tmrLoad 
      Left            =   7920
      Top             =   8550
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00DBE6E6&
      Caption         =   "불러오기"
      Height          =   495
      Left            =   9630
      Style           =   1  '그래픽
      TabIndex        =   26
      Top             =   8550
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBE6E6&
      Height          =   1290
      Left            =   6540
      TabIndex        =   4
      Top             =   390
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
      Left            =   13905
      Style           =   1  '그래픽
      TabIndex        =   3
      Top             =   8550
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "화면지움(&C)"
      Height          =   495
      Left            =   12690
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   8550
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      Height          =   495
      Left            =   11460
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   8550
      Width           =   1215
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   6540
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
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
   Begin MSCommLib.MSComm MSComm 
      Left            =   6570
      Top             =   8415
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MDIControls.MDIActiveX MDIActiveX 
      Left            =   7230
      Top             =   8490
      _ExtentX        =   847
      _ExtentY        =   794
   End
   Begin FPSpread.vaSpread tblReady 
      Height          =   3555
      Left            =   90
      TabIndex        =   21
      Top             =   510
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
      SpreadDesigner  =   "frmIISBactAlert.frx":0000
   End
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4410
      Left            =   90
      TabIndex        =   22
      Top             =   4650
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
      SpreadDesigner  =   "frmIISBactAlert.frx":0509
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   90
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   90
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
      Left            =   90
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4230
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
      Left            =   6540
      TabIndex        =   25
      Top             =   1695
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
      SpreadDesigner  =   "frmIISBactAlert.frx":0D59
      TextTip         =   2
   End
End
Attribute VB_Name = "frmIISBactAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISMGIT960.frm
'   작성자  : 오세원
'   내  용  : MGIT960 장비폼
'   작성일  : 2010-03-23
'   버  전  :
'   병  원  :
'       1. 단국대병원
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady의 Column Enum
'Private Enum TReadyEnum
'    ccNo = 1
'    ccBarNo = 2
'    ccAccNo = 3
'    ccPtId = 4
'    ccName = 5
'End Enum

'## tblReady의 Column Enum
Private Enum TReadyEnum
    ccOrdChk = 1
    ccBarNo = 2
    ccAccNo = 3
    ccPtId = 4
    ccName = 5
    ccCount = 6
    ccSndCnt = 7
    ccRcvCnt = 8
    ccIntBase = 9
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

Dim BlnStic As Boolean
Dim strBuffer As String

Dim strLoadTime As String


Public Property Let EqpCd(ByVal vData As String)
    mEqpCd = vData
End Property

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property

Private Sub cmdLoad_Click()

Dim Rs          As ADODB.Recordset
Dim strSQL      As String
Dim strQryDt    As String
Dim Buffer      As String
Dim strSpcNo    As String

    strQryDt = CStr(DateAdd("d", -7, Now))
    strQryDt = Format(strQryDt, "yyyymmdd")
    
             strSQL = "select distinct SPCNO, TRANSDT, TRANSTM, SPCPOS from ACC203"
    strSQL = strSQL & " where eqpcd = '" & mEqpCd & "'"
    strSQL = strSQL & "   and transdt >= '" & strQryDt & "'"
    
    Set Rs = CliCon.Execute(strSQL, , adCmdText)
    If Not (Rs.BOF Or Rs.EOF) Then
        Do Until Rs.EOF
            If Trim(Rs.Fields("SPCNO").Value) <> "" Then
                If strSpcNo <> Trim(Rs.Fields("SPCNO").Value) Then
                             
                             Buffer = "1H|\^&|||BACT/ALERT^A.00|||||||P|1|20100611104944" & vbCr
                    Buffer = Buffer & "E5" & vbCr
                    Buffer = Buffer & "2P|1|" & vbCr
                    Buffer = Buffer & "BB" & vbCr
                    Buffer = Buffer & "3O|1|0" & Trim(Rs.Fields("SPCNO").Value) & "|0" & Trim(Rs.Fields("SPCNO").Value) & "*19486bb5||||||||||||||||||||||I" & vbCr
                    Buffer = Buffer & "0B" & vbCr
                    Buffer = Buffer & "4R|1|^^^BC^BFA^SFJBTVSH|-|||||I|||" & Trim(Rs.Fields("TRANSDT").Value) & Trim(Rs.Fields("TRANSTM").Value) & "|" & Format(Now, "yyyymmddhhmmss") & "|" & Trim(Rs.Fields("SPCPOS").Value) & vbCr
                    Buffer = Buffer & "67" & vbCr
                    Buffer = Buffer & "5L|1|F" & vbCr
                    Buffer = Buffer & "00" & vbCr
                    Buffer = Buffer & "" & vbCr
        
                    BlnStic = True
                    strBuffer = Buffer
                    mIntLib.Phase = 2
                    
                    Call MSComm_OnComm
                
                    BlnStic = False
                    strBuffer = ""
                    mIntLib.Phase = 1
                    strSpcNo = Trim(Rs.Fields("SPCNO").Value)
                End If
                
            End If
            Rs.MoveNext
        Loop
    End If
    Rs.Close
    Set Rs = Nothing
    
    
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
    
    BlnStic = False
    strBuffer = ""
    
    strLoadTime = Format(Now, "yyyymmddhhmmss")
    
    tmrLoad.Interval = 60000
    tmrLoad.Enabled = True
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Deactivate()
    Me.MDIActiveX.WindowState = ccMinimize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mOrder = Nothing
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISBactAlert = Nothing
End Sub

Private Sub cmdAlarm_Click()
    Dim objErrorShow As clsIISErrorShow     '에러폼 표시 클래스

'Call MSComm_OnComm

'Exit Sub

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

    If BlnStic = True Then
        GoTo Rst
    End If
    
    Select Case MSComm.CommEvent
        Case comEvReceive
            Dim Buffer      As Variant
            Dim BufChar     As String
            Dim lngBufLen   As Long
            Dim i           As Long
Rst:
            If BlnStic = False Then
                Buffer = MSComm.Input
            Else
                Buffer = strBuffer
            End If
            Call mIntLib.WriteLog(Buffer, ccEqp)
    

'             Buffer = ENQ & "1H|\^&|||BACT/ALERT^A.00|||||||P|1|20081022111711" & vbCr
'    Buffer = Buffer & "DF" & vbCr
'    Buffer = Buffer & "2P|1|" & vbCr
'    Buffer = Buffer & "BB" & vbCr
'    Buffer = Buffer & "3O|1|010000191852|010000191852*163556a4||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "F9" & vbCr
'    Buffer = Buffer & "4R|1|^^^BC^BFA^SFH5RSKD|*|||||I|||20081022102623|20081022102623|1D26" & vbCr
'    Buffer = Buffer & "4B" & vbCr
'    Buffer = Buffer & "5R|2|^^^BC^BFN^SGGQBB13|*|||||I|||20081022102619|20081022102619|1D36" & vbCr
'    Buffer = Buffer & "35" & vbCr

'             Buffer = ENQ & "1H|\^&|||BACT/ALERT^A.00|||||||P|1|20100611104944" & vbCr
'    Buffer = Buffer & "E5" & vbCr
'    Buffer = Buffer & "2P|1|" & vbCr
'    Buffer = Buffer & "BB" & vbCr
'    Buffer = Buffer & "3O|1|010000198271|010000198271*19486bb5||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "0B" & vbCr
'    Buffer = Buffer & "4R|1|^^^BC^BFA^SFJBTVSH|-|||||I|||20100611104925|20100615104925|1A01" & vbCr
'    Buffer = Buffer & "67" & vbCr
'    Buffer = Buffer & "5L|1|F" & vbCr
'    Buffer = Buffer & "00" & vbCr
'    Buffer = Buffer & "" & vbCr

'    Buffer = Buffer & "6O|2|180006251047|180006251047*1635567f||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "F1" & vbCr
'    Buffer = Buffer & "7R|1|^^^BC^BFA^SFH5S6FS|*|||||I|||20081022102547|20081022102547|1A24" & vbCr
'    Buffer = Buffer & "41" & vbCr
'    Buffer = Buffer & "0R|2|^^^BC^BFN^SGGQBBBW|*|||||I|||20081022102542|20081022102542|1A34" & vbCr
'    Buffer = Buffer & "56" & vbCr
'    Buffer = Buffer & "1O|3|180006138034|180006138034*162e8619||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "EC" & vbCr
'    Buffer = Buffer & "2R|1|^^^BC^BFN^SGGMRZK9|-|||||P|||20081017062302|20081022063004|1A43" & vbCr
'    Buffer = Buffer & "72" & vbCr
'    Buffer = Buffer & "3R|2|^^^BC^BFA^SFH3FPYH|-|||||P|||20081017062257|20081022063004|1A53" & vbCr
'    Buffer = Buffer & "5E" & vbCr
'    Buffer = Buffer & "4O|4|180006244933|180006244933*163480e5||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "F6" & vbCr
'    Buffer = Buffer & "5R|1|^^^BC^BFA^SFH5RSDY|*|||||I|||20081021191426|20081021191426|2C04" & vbCr
'    Buffer = Buffer & "66" & vbCr
'    Buffer = Buffer & "6R|2|^^^BC^BFN^SGGQBB79|*|||||I|||20081021191420|20081021191420|2C14" & vbCr
'    Buffer = Buffer & "38" & vbCr
'    Buffer = Buffer & "7O|5|180006247118|180006247118*1634c494||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "F8" & vbCr
'    Buffer = Buffer & "0R|1|^^^BC^BFN^SGGQBB7S|*|||||I|||20081022000316|20081022000316|1D01" & vbCr
'    Buffer = Buffer & "3B" & vbCr
'    Buffer = Buffer & "1R|2|^^^BC^BFA^SFH5RSF4|*|||||I|||20081022000307|20081022000307|1D12" & vbCr
'    Buffer = Buffer & "27" & vbCr
'    Buffer = Buffer & "2O|6|180006140587|180006140587*162e8603||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "F5" & vbCr
'    Buffer = Buffer & "3R|1|^^^BC^BFA^SFH5S4X2|-|||||P|||20081017062239|20081022063004|1A55" & vbCr
'    Buffer = Buffer & "3B" & vbCr
'    Buffer = Buffer & "4R|2|^^^BC^BFN^SGGQB9R3|-|||||P|||20081017062234|20081022063004|1A45" & vbCr
'    Buffer = Buffer & "4F" & vbCr
'    Buffer = Buffer & "5O|7|180006244681|180006244681*163480f2||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "F8" & vbCr
'    Buffer = Buffer & "6R|1|^^^BC^BFA^SFH5S8C6|*|||||I|||20081021191437|20081021191437|2C05" & vbCr
'    Buffer = Buffer & "2E" & vbCr
'    Buffer = Buffer & "7R|2|^^^BC^BFN^SGGQBB76|*|||||I|||20081021191433|20081021191433|2C15" & vbCr
'    Buffer = Buffer & "3F" & vbCr
'    Buffer = Buffer & "0O|8|180006244810|180006244810*163480d9||||||||||||||||||||||P" & vbCr
'    Buffer = Buffer & "F4" & vbCr
'    Buffer = Buffer & "1R|1|^^^BC^BFA^SFH5RSCD|*|||||I|||20081021191415|20081021191415|2C03" & vbCr
'    Buffer = Buffer & "47" & vbCr
'    Buffer = Buffer & "2R|2|^^^BC^BFN^SGGQBB91|+|||||P|||20081021191409|20081022082236|2C13" & vbCr
'    Buffer = Buffer & "41" & vbCr
'    Buffer = Buffer & "3R|3|^^^TTD^BFN^SGGQBB91|0013.1|||||P|||20081021191409|20081022082236|2C13" & vbCr
'    Buffer = Buffer & "A2" & vbCr
'    Buffer = Buffer & "4O|9|180006250286|180006250286*163553e3||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "F6" & vbCr
'    Buffer = Buffer & "5R|1|^^^BC^BFN^SGGQB9PN|*|||||I|||20081022101439|20081022101439|2C52" & vbCr
'    Buffer = Buffer & "61" & vbCr
'    Buffer = Buffer & "6R|2|^^^BC^BFA^SFH5S5D2|*|||||I|||20081022101434|20081022101434|2C42" & vbCr
'    Buffer = Buffer & "14" & vbCr
'    Buffer = Buffer & "7O|10|180006228247|180006228247*163456a1||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "21" & vbCr
'    Buffer = Buffer & "0R|1|^^^BC^BFN^SGGQ9KMZ|*|||||I|||20081021161405|20081022103916|2D40" & vbCr
'    Buffer = Buffer & "6C" & vbCr
'    Buffer = Buffer & "1R|2|^^^BC^BFA^SFH5S6BP|*|||||I|||20081021161400|20081021161400|2D30" & vbCr
'    Buffer = Buffer & "26" & vbCr
'    Buffer = Buffer & "2O|11|180006246302|180006246302*1634b089||||||||||||||||||||||I" & vbCr
'    Buffer = Buffer & "13" & vbCr
'    Buffer = Buffer & "3R|1|^^^BC^BFA^SFH5RPGW|*|||||I|||20081021223741|20081021223741|2C24" & vbCr
'    Buffer = Buffer & "5C" & vbCr
'    Buffer = Buffer & "4L|1" & vbCr
'    Buffer = Buffer & "3D" & vbCr
'    Buffer = Buffer & "" & vbCr
    
'             Buffer = ENQ & "1H|\^&|||Becton Dickinson||||||||V1.0|20100323193220" & vbCr
'    Buffer = Buffer & "P|1" & vbCr
'    Buffer = Buffer & "O|1|100001738671||^^^MGIT_960_GND" & vbCr
'    Buffer = Buffer & "R|1|^^^GND_MGIT^430158909489|INST_ONGOING^0|||||P|||20100323193142||MGIT960^^45^1^A/F03" & vbCr
'    Buffer = Buffer & "L|1|N" & vbCr
'    Buffer = Buffer & "2C" & vbCr
'    Buffer = Buffer & "" & vbCr
    
    
'             Buffer = ENQ & "1H|\^&|||Becton Dickinson||||||||V1.0|20100507010236" & vbCr
'    Buffer = Buffer & "P|1" & vbCr
'    Buffer = Buffer & "O|1|100003557872||^^^MGIT_960_GND" & vbCr
'    Buffer = Buffer & "R|1|^^^GND_MGIT^430158901934|INST_ONGOING^78|||||P|||20100504150604|20100507010232|MGIT960^^42^1^A/J19" & vbCr
'    Buffer = Buffer & "L|1|N" & vbCr
'    Buffer = Buffer & "6F" & vbCr4
'    Buffer = Buffer & "" & vbCr

'             Buffer = ENQ & "
'1H|\^&|||Becton Dickinson||||||||V1.0|20100507200236
'P|1
'O|1|100001738671||^^^MGIT_960_GND
'R|1|^^^GND_MGIT^430158909489|INST_NEGATIVE^0|||||P|||20100323193420|20100507193420|MGIT960^^45^1^A/F04
'L|1|N
'30
'
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
'                                mIntLib.BufCnt = mIntLib.BufCnt + 1
                                mIntLib.Phase = 3
                            Case vbCr, vbLf
                                mIntLib.BufCnt = mIntLib.BufCnt + 1
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
    Dim strResult    As String   '수신한 결과
    Dim strFlag      As String   '수신한 Abnormal Flag
    Dim strComm      As String   '수신한 Comment
    Dim strTemp1     As String
    Dim strTemp2     As String
    Dim strOrdTime   As String
    Dim strRstTime   As String
    
    
    Set objIntNms = mIntLib.IntNms
    For Each objBuffer In mIntLib.Buffers
        strRcvBuf = objBuffer.Buffers
        strType = Mid$(strRcvBuf, 2, 1)
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "O"    '## Order
                strBarNo = Format$(mGetP(strRcvBuf, 3, "|"), String$(SPCLEN, "#"))
                Call GetOrder(strBarNo)
            Case "R"    '## Result
                '## 장비기준 검사명, 결과, Abnormal Flag
                strTemp1 = mGetP(strRcvBuf, 4, "|")                     '-- 결과
                strTemp2 = mGetP(strRcvBuf, 14, "|")                    '-- 포지션
                strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")    '-- Channel??
                strRackNo = strTemp2
                
                strOrdTime = mGetP(strRcvBuf, 12, "|")
                strRstTime = mGetP(strRcvBuf, 13, "|")
                
                '-- 결과유형 : *,-,+
                If strTemp1 <> "" Then
                    Select Case strTemp1
                        Case "*":   strResult = "*" '-- 검사시작
                        Case "-":   strResult = "-"
                        Case "+":   strResult = "+"
                        Case Else:  strResult = strTemp1
                    End Select
                    
                    If strResult = "-" Then
                        '-- 장비에서 나오는 값은 최종값만 있음
'                        If CStr(DateAdd("d", 2, Format(Mid(strOrdTime, 1, 8), "####-##-##"))) = CStr(Format(Mid(strRstTime, 1, 8), "####-##-##")) Then
'                            strResult = "C3"     '접수일로부터 2일
'                        ElseIf CStr(DateAdd("d", 5, Format(Mid(strOrdTime, 1, 8), "####-##-##"))) = CStr(Format(Mid(strRstTime, 1, 8), "####-##-##")) Then
'                            strResult = "C17"     '-- 최종
'                        ElseIf CStr(DateAdd("d", 7, Format(Mid(strOrdTime, 1, 8), "####-##-##"))) = CStr(Format(Mid(strRstTime, 1, 8), "####-##-##")) Then
'                            strResult = "C4"     '-- 최종
'                        Else
'                            strResult = ""
'                        End If
                    
                    
'                        If CStr(DateAdd("h", 48, Format(strOrdTime, "####-##-## ##:##:##"))) = CStr(Format(strRstTime, "####-##-## ##:##:##")) Then '-- 2day no growth
'                            strResult = "C3"
'                        ElseIf CStr(DateAdd("h", 120, Format(strOrdTime, "####-##-## ##:##:##"))) = CStr(Format(strRstTime, "####-##-## ##:##:##")) Then '-- 5day no growth
'                            strResult = "C17"
'                        ElseIf CStr(DateAdd("h", 168, Format(strOrdTime, "####-##-## ##:##:##"))) = CStr(Format(strRstTime, "####-##-## ##:##:##")) Then '-- 7day no growth
'                            strResult = "C4"     '-- 최종
'                        Else
'                            strResult = ""
'                        End If
                                        
                        If CStr(DateAdd("h", 48, Format(strOrdTime, "####-##-## ##:##:##"))) = CStr(DateAdd("h", 0, Format(strRstTime, "####-##-## ##:##:##"))) Then '-- 2day no growth
                            strResult = "C3"
                        ElseIf CStr(DateAdd("h", 120, Format(strOrdTime, "####-##-## ##:##:##"))) = CStr(DateAdd("h", 0, Format(strRstTime, "####-##-## ##:##:##"))) Then '-- 5day no growth
                            strResult = "C17"
                        ElseIf CStr(DateAdd("h", 168, Format(strOrdTime, "####-##-## ##:##:##"))) = CStr(DateAdd("h", 0, Format(strRstTime, "####-##-## ##:##:##"))) Then '-- 7day no growth
                            strResult = "C4"     '-- 최종
                        Else
                            strResult = ""
                        End If
                    ElseIf strResult = "+" Then
                            strResult = "C12"     '-- 균 동정중 ==> 결과저장 않는다
                    End If
                        
                    Set objIntInfo = New clsIISIntInfo
                    With objIntInfo
                        .BarNo = strBarNo
                        .SpcPos = strRackNo
                        '-- 결과값을 넣을곳이 마땅치 않음..
                        .VitekNo = strResult
                    End With
                    
                    '## 정성결과 저장
                    If objIntNms.ExistIntBase(strIntBase) Then
                        Call objIntInfo.IntResults.Add(strIntBase, objIntNms.GetIntNm(strIntBase), _
                             strTemp1, strResult, strRackNo)
                    End If
                    mIntLib.State = "R"
                End If
                
            Case "C"    '## Comment
            
            Case "L"    '## Terminator
                '## DB에 결과저장
                If mIntLib.State = "R" Then
                    Call SaveServer(objIntInfo, LCase(strTemp1), strResult, strIntBase)
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
Private Sub SaveServer(ByVal pIntInfo As clsIISIntInfo, Optional ByVal pIntResult As String, Optional pResult As String, Optional pIntBase As String)
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

        Call SetComplete2(objAccInfo, pIntResult, pResult, pIntBase)
        Call tblComplete_Click(1, tblComplete.DataRowCnt)
        Set objAccInfo = Nothing
        
        '## ClientDb, Server에 결과저장
        Call mIntLib.SaveClientDb(strSpcYy, lngSpcNo)
        
        '## 결과저장
        Call mIntLib.SaveCultureResult(pIntInfo)
        
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

''-----------------------------------------------------------------------------'
''   기능 : 해당 바코드번호에 대한 접수정보 조회, tblReady, tblResult에 표시
''   인수 :
''       - pBarNo : 바코드번호
''-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim strOutput  As String            '송신할 데이터

    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    If Not (objAccInfo Is Nothing) Then
        '## tblReady, tblresult, Label에 정보표시
        Call SetReady(objAccInfo)
        Call SetLabel(objAccInfo)
        Call SetResult(objAccInfo)

'        mOrder.AccInfo = objAccInfo
'        mOrder.NoOrder = False
        Set objAccInfo = Nothing
    Else
'        mOrder.NoOrder = True
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
'Private Sub SendOrder()
'    Dim strOutput As String     '송신할 데이터
'
'    Select Case mIntLib.SndPhase
'        Case 1  '## Header
'            strOutput = mIntLib.GetFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
'            mIntLib.SndPhase = 2
'
'        Case 2  '## Patient
'            strOutput = mIntLib.GetFrameNo & "P|1" & vbCr & ETX
'            mIntLib.SndPhase = 4
'
'        Case 3  '## No Order
'
'        Case 4  '## Order
'            With mOrder
'                If .NoOrder = True Then
'                    '## 접수정보가 없을경우
'                    strOutput = mIntLib.GetFrameNo & "O|1|" & .BarNo & "|" & .Seq & "^" & .RackNo & _
'                                "^" & .TubePos & "^^SAMPLE^NORMAL|ALL" & _
'                                "|R||||||C||||||||||||||Q" & vbCr & ETX
'                    mIntLib.SndPhase = 5
'                Else
'                    If .IsSending = False Then  '## 최초 보낼때
'                        strOutput = "O|1|" & .BarNo & "|" & .Seq & "^" & .RackNo & "^" & .TubePos & _
'                                    "^^SAMPLE^NORMAL|" & .GetOrder & "|R||||||N||||||||||||||Q"
'                        If Len(strOutput) > 230 Then
'                            .IsSending = True
'                            .Order = Mid$(strOutput, 231)
'                            strOutput = mIntLib.GetFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                            mIntLib.SndPhase = 4
'                        Else
'                            strOutput = mIntLib.GetFrameNo & strOutput & vbCr & ETX
'                            mIntLib.SndPhase = 5
'                        End If
'                    Else                        '## 남은 문자열이 있을때
'                        strOutput = .Order
'                        If Len(strOutput) > 230 Then
'                            .Order = Mid$(strOutput, 231)
'                            strOutput = mIntLib.GetFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                            mIntLib.SndPhase = 4
'                        Else
'                            .IsSending = False
'                            strOutput = mIntLib.GetFrameNo & strOutput & vbCr & ETX
'                            mIntLib.SndPhase = 5
'                        End If
'                    End If
'                End If
'            End With
'
'        Case 5  '## Termianator
'            strOutput = mIntLib.GetFrameNo & "L|1" & vbCr & ETX
'            mIntLib.SndPhase = 6
'
'        Case 6  '## EOT
'            mIntLib.State = ""
'            MSComm.Output = EOT
'            Call mIntLib.WriteLog(EOT, ccPCLog)
'            Exit Sub
'    End Select
'
'    strOutput = STX & strOutput & mOrder.GetChkSum(strOutput) & vbCrLf
'    MSComm.Output = strOutput
'    Call mIntLib.WriteLog(strOutput, ccPCLog)
'End Sub

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

'        .Col = TReadyEnum.ccNo:     .Text = mOrder.TubePos & "/" & Mid$(mOrder.RackNo, 2)
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
    Dim intRow       As Integer
    
    With tblComplete
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = TCompleteEnum.ccBarNo
            If Trim(.Text) = Trim(pIntInfo.BarNo) Then
                Exit Sub
            End If
        Next
        
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
Private Sub SetComplete2(ByVal pAccInfo As clsIISAccInfo, Optional ByVal pIntResult As String, Optional pResult As String, Optional pIntBase As String)
    Dim objResult   As clsIISResult     '결과내역 클래스
    Dim objQCResult As clsIISQCResult   'QC결과내역 클래스
    Dim vBarNo      As Variant          'Spread의 바코드번호
    Dim vTubePos    As Variant          'Spread의 Tube Position
    Dim i           As Long
    Dim intRow As Integer
    
'    Set objResult = New clsIISResult      '결과내역 클래스

    With tblComplete
        For intRow = 1 To .MaxRows
            .Row = intRow
            .Col = TReadyEnum.ccBarNo
            If Trim(.Text) = Trim(pAccInfo.GetBarNo) And pResult = "" Then
                Exit Sub
            End If
        Next
        
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
            
            If .MaxCols <= .DataColCnt Then
                .MaxCols = .MaxCols + 1
            End If

            .Col = TCompleteEnum.ccResult
            .ColHidden = True
            .Text = pIntBase & DIV & pIntResult & DIV & pResult & _
                    DIV & "" & DIV & "" & DIV & "" & _
                    DIV & "" & DIV & ""
                
'            For Each objResult In pAccInfo.Results
'                If .MaxCols <= .DataColCnt Then
'                    .MaxCols = .MaxCols + 1
'                End If
'
'                .Col = TCompleteEnum.ccResult + i
'                .ColHidden = True
'                .Text = objResult.IntNm.IntNm & DIV & objResult.IntResult & DIV & objResult.RstCd & _
'                        DIV & objResult.Unit & DIV & objResult.HLDiv & DIV & objResult.DPDiv & _
'                        DIV & mGetRef(objResult.Ref.RefFrVal, objResult.Ref.RefToVal) & DIV & _
'                        objResult.IntInfo
'                i = i + 1
'            Next
'            Set objResult = Nothing
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
    If Err.Number = 8005 Then
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


Private Sub tmrLoad_Timer()

Dim strRstTime As String

    strRstTime = Format(Now, "yyyymmddhhmmss")

    If chkLoad.Value = "1" Then
        If CStr(DateAdd("h", 1, Format(strLoadTime, "####-##-## ##:##:##"))) <= CStr(DateAdd("h", 0, Format(strRstTime, "####-##-## ##:##:##"))) Then
            Call cmdLoad_Click
        End If
    End If

End Sub

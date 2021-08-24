VERSION 5.00
Object = "{C8094403-41E7-4EF2-826E-2A56177BCC48}#1.0#0"; "MDIControls.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmIISVitekII 
   BackColor       =   &H00DBE6E6&
   Caption         =   "Vitek II"
   ClientHeight    =   9180
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
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
   LockControls    =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows 기본값
   Begin FPSpread.vaSpread tblSensi 
      Height          =   5835
      Left            =   6555
      TabIndex        =   13
      Top             =   2505
      Width           =   8580
      _Version        =   393216
      _ExtentX        =   15134
      _ExtentY        =   10292
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
      MaxCols         =   3
      MaxRows         =   19
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISVitekII.frx":0000
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
      Left            =   1440
      TabIndex        =   0
      Text            =   "123456789011"
      Top             =   512
      Width           =   1530
   End
   Begin VB.TextBox txtWorkarea 
      Alignment       =   2  '가운데 맞춤
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1515
      MaxLength       =   2
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txtAccDt 
      Alignment       =   2  '가운데 맞춤
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1935
      MaxLength       =   4
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtAccSeq 
      Alignment       =   2  '가운데 맞춤
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2490
      MaxLength       =   4
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CheckBox chkAccNo 
      BackColor       =   &H00DBE6E6&
      Caption         =   "접수번호"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3165
      TabIndex        =   1
      Top             =   570
      Width           =   1035
   End
   Begin FPSpread.vaSpread tblTemp 
      Height          =   3420
      Left            =   8310
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   5940
      _Version        =   393216
      _ExtentX        =   10478
      _ExtentY        =   6033
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
      MaxCols         =   4
      SpreadDesigner  =   "frmIISVitekII.frx":045A
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00DBE6E6&
      Caption         =   "Send(&S)"
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
      Left            =   10290
      Style           =   1  '그래픽
      TabIndex        =   5
      Top             =   8567
      Width           =   1215
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   6548
      TabIndex        =   9
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
      Caption         =   "■ 균정보"
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   8
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
      Height          =   3270
      Left            =   105
      TabIndex        =   10
      Top             =   855
      Width           =   6345
      _Version        =   393216
      _ExtentX        =   11192
      _ExtentY        =   5768
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
      MaxRows         =   10
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISVitekII.frx":1CBC
      TextTip         =   2
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   98
      TabIndex        =   11
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
      TabIndex        =   12
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
   Begin FPSpread.vaSpread tblComplete 
      Height          =   4410
      Left            =   105
      TabIndex        =   15
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
      SpreadDesigner  =   "frmIISVitekII.frx":234F
      TextTip         =   2
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   375
      Left            =   6555
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2115
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
      Caption         =   "■ 항생제 정보"
      Appearance      =   0
   End
   Begin FPSpread.vaSpread tblMnm 
      Height          =   1560
      Left            =   6555
      TabIndex        =   17
      Top             =   495
      Width           =   8580
      _Version        =   393216
      _ExtentX        =   15134
      _ExtentY        =   2752
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
      MaxCols         =   3
      MaxRows         =   4
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13697023
      SpreadDesigner  =   "frmIISVitekII.frx":29E0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "개수 :"
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
      Left            =   4770
      TabIndex        =   24
      Top             =   4380
      Width           =   540
   End
   Begin VB.Label lblCompleteCnt 
      BackStyle       =   0  '투명
      Caption         =   "100"
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
      Left            =   5430
      TabIndex        =   23
      Top             =   4380
      Width           =   450
   End
   Begin VB.Label lblReadyCnt 
      BackStyle       =   0  '투명
      Caption         =   "100"
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
      Left            =   5430
      TabIndex        =   22
      Top             =   585
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "개수 :"
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
      Left            =   4770
      TabIndex        =   21
      Top             =   585
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000000&
      FillColor       =   &H80000005&
      Height          =   315
      Left            =   1440
      Top             =   510
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   180
      Left            =   2385
      TabIndex        =   20
      Top             =   585
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   180
      Left            =   1770
      TabIndex        =   19
      Top             =   585
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblNo 
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
      Left            =   210
      TabIndex        =   14
      Top             =   585
      Width           =   1200
   End
End
Attribute VB_Name = "frmIISVitekII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISVitekII.frm
'   작성자  : 이상대
'   내  용  : Vitek II 장비폼
'   작성일  : 2005-01-31
'   버  전  :
'       1. 1.0.1: 이상대(2005-02-25)
'   병  원  :
'       1. 예수병원
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady의 Column Enum
Private Enum TReadyEnum
    ccOrdChk = 1:       ccNo = 2
    ccBarNo = 3:        ccAccNo = 4
    ccPtId = 5:         ccName = 6
    ccTestNm = 7:       ccSpcNm = 8
    ccSex = 9:          ccAge = 10
    ccDoctNm = 11:      ccDeptNm = 12
    ccWardNm = 13:      ccWorkSheet = 14
End Enum

'## tblComplete의 Column Enum
Private Enum TCompleteEnum
    ccNo = 1:           ccBarNo = 2
    ccAccNo = 3:        ccPtId = 4
    ccName = 5:         ccTestNm = 6
    ccSpcNm = 7:        ccSex = 8
    ccAge = 9:          ccDoctNm = 10
    ccDeptNm = 11:      ccWardNm = 12
    ccWorkSheet = 13
End Enum

'## tblTemp의 Column Enum
Private Enum TTempEnum
    ccNo = 1
    ccMnmCd = 2
    ccMnmNm = 3
    ccCount = 4
    ccResult = 5
End Enum

''## tblMnm의 Column Enum
Private Enum TMnmEnum
    ccMnmCd = 1
    ccMnmNm = 2
    ccNo = 3
End Enum

'## tblSensi Column Enum
Private Enum TSensiEnum
    ccDrugNm = 1
    ccRstCd = 2
    ccVolumn = 3
End Enum

'## Datalog Field 상수
Private Const RS As String = ""    'Record Separator
Private Const GS As String = ""    'Group Separator
Private Const FS As String = "|"    'Field Separator

'## Popup Menu ID
Private Const DELETE    As Long = 1
Private Const DELETEALL As Long = 2

Private WithEvents mIntLib  As clsIISInterface   '인터페이스 클래스
Attribute mIntLib.VB_VarHelpID = -1
Private WithEvents mPopup   As clsIISPopup       '팝업메뉴
Attribute mPopup.VB_VarHelpID = -1

Private mIntErrors  As clsIISIntErrors           '인터페이스 에러 컬렉션
Private mOrder      As clsIISIntOrder            '오더정보 클래스

Private mEqpCd  As String   '장비코드
Private mEqpKey As String   '장비키

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
    Set mOrder = New clsIISIntOrder
    Set mIntLib = New clsIISInterface

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
    Set mIntLib = Nothing
    Set mIntErrors = Nothing
    Set frmIISVitekII = Nothing
End Sub

Private Sub cmdAlarm_Click()
    Dim objErrorShow As clsIISErrorShow     '에러폼 표시 클래스

    Set objErrorShow = New clsIISErrorShow
    Call objErrorShow.ShowErrors(mIntErrors)
    Set objErrorShow = Nothing

    '## 에러가 없으면 버튼색깔 원래대로, 있으면 계속 빨강색
    cmdAlarm.BackColor = IIf(mIntErrors.Count = 0, &HF4F0F2, vbRed)
    
    '## 1.0.1: 이상대(2005-02-25)
    '   - Alarm창이 닫힌후 포커스를 txtBarNo Or txtAccSeq로 이동
    If chkAccNo.Value = 1 Then
        txtAccSeq.SetFocus
    Else
        txtBarNo.SetFocus
    End If
End Sub

Private Sub cmdClear_Click()
    Call CtlClear
    Call mIntLib.AccInfos.RemoveAll

    If chkAccNo.Value = 1 Then
        Call SetAccNo
        txtAccSeq.SetFocus
    Else
        txtBarNo.SetFocus
    End If
End Sub

Private Sub cmdSend_Click()
    Dim vOrdChk As Variant  'Spread의 오더전송여부
    Dim i       As Long
    
    '## 포트가 오픈되어 있지 않으면 에러표시
    If MSComm.PortOpen = False Then
        MsgBox "포트가 열려있지 않습니다.", vbCritical, "오류"
        Exit Sub
    End If

    With tblReady
        If .DataRowCnt < 1 Then Exit Sub
        
        '## 송신할 검체개수 파악!
        mOrder.SendCnt = 0
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccOrdChk, i, vOrdChk)
            
            If CStr(vOrdChk) = "" Then
                mOrder.SendCnt = mOrder.SendCnt + 1
            End If
        Next i
    End With
        
    '## ENQ 전송
    MSComm.Output = ENQ
    Call mIntLib.WriteLog(ENQ, ccPCLog)
    mIntLib.State = "Q"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub chkAccNo_Click()
    If chkAccNo.Value = 1 Then
        lblNo.Caption = "접수번호 : "
        txtBarNo.Visible = False
        Shape1.Visible = True:      txtWorkarea.Visible = True:
        txtAccDt.Visible = True:    txtAccSeq.Visible = True:
        Label1.Visible = True:      Label1.ZOrder 0
        Label2.Visible = True:      Label2.ZOrder 0
        Call SetAccNo:              txtAccSeq.SetFocus
    Else
        lblNo.Caption = "바코드번호 : "
        txtBarNo.Visible = True
        Shape1.Visible = False:     txtWorkarea.Visible = False:
        txtAccDt.Visible = False:   txtAccSeq.Visible = False:
        txtBarNo.SetFocus
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
        lblReadyCnt.Caption = CStr(tblReady.DataRowCnt)
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

Private Sub txtWorkarea_GotFocus()
    With txtWorkarea
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtWorkarea_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtAccDt_GotFocus()
    With txtAccDt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAccDt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtAccSeq_GotFocus()
    With txtAccSeq
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtAccSeq_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtWorkarea.Text = "" Then Exit Sub
    If txtAccDt.Text = "" Then Exit Sub
    If txtAccSeq.Text = "" Then Exit Sub
    
    If KeyCode = vbKeyReturn Then
        Me.MousePointer = vbHourglass
        
        '## 2100년 부터는 이프로그램 사용못함!
        Call GetOrderX(Trim$(txtWorkarea.Text), "20" & Trim$(txtAccDt.Text), Trim$(txtAccSeq.Text))
        lblReadyCnt.Caption = CStr(tblReady.DataRowCnt)
        Me.MousePointer = vbDefault
    End If
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

Private Sub tblReady_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim vBarNo      As Variant          'Spread의 바코드번호
    Dim strInfo     As String           '접수정보
    
    If Row = 0 Then Exit Sub
    With tblReady
        Call .GetText(TReadyEnum.ccBarNo, Row, vBarNo)
        If Trim$(CStr(vBarNo)) = "" Then Exit Sub
        
        .Row = Row
        strInfo = vbCrLf
        .Col = TReadyEnum.ccTestNm
        strInfo = strInfo & Space(2) & "검 사 명 : " & .Text & vbCrLf
        .Col = TReadyEnum.ccSpcNm
        strInfo = strInfo & Space(2) & "검 체 명 : " & .Text & vbCrLf
        .Col = TReadyEnum.ccSex
        strInfo = strInfo & Space(2) & "성    별 : " & .Text & vbCrLf
        .Col = TReadyEnum.ccAge
        strInfo = strInfo & Space(2) & "나    이 : " & .Text & vbCrLf
        .Col = TReadyEnum.ccDoctNm
        strInfo = strInfo & Space(2) & "처 방 의 : " & .Text & vbCrLf
        .Col = TReadyEnum.ccDeptNm
        strInfo = strInfo & Space(2) & "진 료 과 : " & .Text & vbCrLf
        .Col = TReadyEnum.ccWardNm
        strInfo = strInfo & Space(2) & "병    동 : " & .Text & vbCrLf
        .Col = TReadyEnum.ccWorkSheet
        strInfo = strInfo & Space(2) & "WorkSheet Unit : " & .Text & vbCrLf
        
        MultiLine = 1
        TipWidth = 6000
        TipText = strInfo
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
End Sub

Private Sub tblComplete_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vVitekNo1   As Variant      'Spread의 Vitek No1(tblComplete)
    Dim vVitekNo2   As Variant      'Spread의 Vitek No2(tblTemp)
    Dim i           As Long
    
    Call mTblClear(tblMnm)
    Call mTblClear(tblSensi)
    
    Call tblComplete.GetText(TCompleteEnum.ccNo, Row, vVitekNo1)
    With tblMnm
        For i = 1 To tblTemp.DataRowCnt
            Call tblTemp.GetText(TTempEnum.ccNo, i, vVitekNo2)
            
            If vVitekNo1 = vVitekNo2 Then
                If .MaxRows <= .DataRowCnt Then
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                Else
                    .Row = .DataRowCnt + 1
                End If
                
                tblTemp.Row = i
                tblTemp.Col = TTempEnum.ccMnmCd
                .Col = TMnmEnum.ccMnmCd:    .Text = tblTemp.Value
                
                tblTemp.Col = TTempEnum.ccMnmNm
                .Col = TMnmEnum.ccMnmNm:    .Text = tblTemp.Value
                .Col = TMnmEnum.ccNo:       .Text = CStr(vVitekNo1)
            End If
        Next i
        
        '## 해당 균코드의 항생제결과 조회
        If .DataRowCnt > 0 Then
            Call .SetActiveCell(1, 1)
            Call tblMnm_Click(1, 1)
        End If
    End With
End Sub

Private Sub tblComplete_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim vBarNo      As Variant          'Spread의 바코드번호
    Dim strInfo     As String           '접수정보
    
    If Row = 0 Then Exit Sub
    With tblComplete
        Call .GetText(TReadyEnum.ccBarNo, Row, vBarNo)
        If Trim$(CStr(vBarNo)) = "" Then Exit Sub
        
        .Row = Row
        strInfo = vbCrLf
        .Col = TCompleteEnum.ccTestNm
        strInfo = strInfo & Space(2) & "검 사 명 : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccSpcNm
        strInfo = strInfo & Space(2) & "검 체 명 : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccSex
        strInfo = strInfo & Space(2) & "성    별 : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccAge
        strInfo = strInfo & Space(2) & "나    이 : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccDoctNm
        strInfo = strInfo & Space(2) & "처 방 의 : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccDeptNm
        strInfo = strInfo & Space(2) & "진 료 과 : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccWardNm
        strInfo = strInfo & Space(2) & "병    동 : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccWorkSheet
        strInfo = strInfo & Space(2) & "WorkSheet Unit : " & .Text & vbCrLf
        
        MultiLine = 1
        TipWidth = 6000
        TipText = strInfo
        Call .SetTextTipAppearance("돋움체", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
End Sub

Private Sub tblMnm_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vVitekNo1   As Variant  'Spread의 Vitek No1(tblMnm)
    Dim vVitekNo2   As Variant  'Spread의 Vitek No2(tblTemp)
    Dim vMnmCd1     As Variant  'Spread의 균코드1(tblMnm)
    Dim vMnmCd2     As Variant  'Spread의 균코드2(tblTemp)
    Dim vCount      As Variant  'Spread의 항생제개수(tblTemp)
    Dim i           As Long
    Dim j           As Long
    
    Call mTblClear(tblSensi)
    
    Call tblMnm.GetText(TMnmEnum.ccNo, Row, vVitekNo1)
    Call tblMnm.GetText(TMnmEnum.ccMnmCd, Row, vMnmCd1)
    With tblSensi
        For i = 1 To tblTemp.DataRowCnt
            Call tblTemp.GetText(TTempEnum.ccNo, i, vVitekNo2)
            Call tblTemp.GetText(TTempEnum.ccMnmCd, i, vMnmCd2)
            
            If vVitekNo1 = vVitekNo2 And vMnmCd1 = vMnmCd2 Then
                Call tblTemp.GetText(TTempEnum.ccCount, i, vCount)
                
                For j = 0 To CLng(vCount) - 1
                    If .MaxRows <= .DataRowCnt Then
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                    Else
                        .Row = .DataRowCnt + 1
                    End If
                    
                    tblTemp.Row = i
                    tblTemp.Col = TTempEnum.ccResult + j
                    .Col = TSensiEnum.ccDrugNm: .Text = mGetP(tblTemp.Text, 2, DIV)
                    .Col = TSensiEnum.ccRstCd:  .Text = mGetP(tblTemp.Text, 3, DIV)
                    .Col = TSensiEnum.ccVolumn: .Text = mGetP(tblTemp.Text, 4, DIV)
                Next j
                Exit For
            End If
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
            Dim lngCheckSum As Long
            Dim i           As Long

            Buffer = MSComm.Input
            Call mIntLib.WriteLog(Buffer, ccEqp)

            lngBufLen = Len(Buffer)
            For i = 1 To lngBufLen
                BufChar = Mid$(Buffer, i, 1)
                
                Select Case mIntLib.Phase
                    Case 1      '## ENQ, ACK 대기
                        Select Case BufChar
                            Case ENQ
                                mIntLib.BufCnt = 1
                                Call mIntLib.ClearBuffer
                                
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                                mIntLib.Phase = 2
                            Case ACK
                                If mIntLib.State = "Q" Then     '## ENQ 전송후
                                    Call SendOrder
                                ElseIf mIntLib.State = "C" Then '## CheckSum 전송후
                                    '## 전송한 검체 Check 표시
                                    Call tblReady.SetText(TReadyEnum.ccOrdChk, mOrder.Seq, "√")
                                    mOrder.SendCnt = mOrder.SendCnt - 1
                                    
                                    '## ETX 전송
                                    MSComm.Output = ETX
                                    Call mIntLib.WriteLog(ETX, ccPCLog)
                                    
                                    '## EOT 전송
                                    MSComm.Output = EOT
                                    Call mIntLib.WriteLog(EOT, ccPCLog)
                                    
                                    '## 전송할 검체가 있으면 ENQ전송
                                    If mOrder.SendCnt > 0 Then
                                        Call mSleep(1000)
                                        mIntLib.State = "Q"
                                        MSComm.Output = ENQ
                                        Call mIntLib.WriteLog(ENQ, ccPCLog)
                                    End If
                                End If
                        End Select
                    Case 2      '## GS 대기
                        Select Case BufChar
                            Case STX
                            Case GS
                                mIntLib.Phase = 3
                            Case Else
                                Call mIntLib.AddBuffer(BufChar)
                        End Select
                    Case 3      '## CheckSum 대기
                        lngCheckSum = lngCheckSum + 1
                        If lngCheckSum = 2 Then
                            MSComm.Output = ACK
                            Call mIntLib.WriteLog(ACK, ccPCLog)
                            mIntLib.Phase = 4
                        End If
                    Case 4      '## CheckSum 대기
                        Select Case BufChar
                            Case ETX
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
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim objAccInfo  As clsIISAccInfo    '접수정보 클래스
    Dim vOrdChk     As Variant          'Spread의 오더전송유무
    Dim vBarNo      As Variant          'Spread의 바코드번호
    Dim vVitekNo    As Variant          'Spread의 VitekNo
    Dim strOutput   As String           '송신할 데이터
    Dim blnSend     As Boolean          '오더전송 유무
    Dim i           As Long
    
    With tblReady
        '## Spread에서 오더를 전송하지 않은 검체를 검색후 오더전송
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccOrdChk, i, vOrdChk)
            
            If CStr(vOrdChk) = "" Then
                Call .GetText(TReadyEnum.ccNo, i, vVitekNo)
                Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
                
                Set objAccInfo = mIntLib.GetAccInfoX(CStr(vBarNo))
                If Not (objAccInfo Is Nothing) Then
                    '## 오더정보 클래스 초기화
                    blnSend = True
                    mOrder.ClsClear
                    mOrder.Seq = i
                    
                    '## 1.STX 전송
                    strOutput = STX & vbCrLf
                    MSComm.Output = strOutput
                    Call mIntLib.WriteLog(strOutput, ccPCLog)
                    
                    '## 2.오더문자열 전송
                    strOutput = mOrder.GetOrder(objAccInfo, CStr(vVitekNo))
                    MSComm.Output = strOutput
                    Call mIntLib.WriteLog(strOutput, ccPCLog)
                    
                    '## 3.CheckSum 전송
                    strOutput = GS & mOrder.CheckSum & vbCrLf
                    MSComm.Output = strOutput
                    Call mIntLib.WriteLog(strOutput, ccPCLog)
                    
                    mIntLib.State = "C"
                    Set objAccInfo = Nothing
                    Exit Sub
                End If
            End If
        Next i
        
        '## 잘못된 오더의 경우 EOT 전송하여 통신종료
        If blnSend = False Then
            mIntLib.State = ""
            MSComm.Output = EOT
            Call mIntLib.WriteLog(EOT, ccPCLog)
        End If
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 장비로부터 수신한 데이터 편집
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim objVitek     As clsIISIntInfo   '장비에서 수신한 결과저장 클래스
    
    Dim aryTemp1()   As String
    Dim aryTemp2()   As String
    Dim strRcvBuf    As String   '수신한 Data
    Dim strCode      As String   '수신한 Field Code
    Dim strDrugCd    As String   '수신한 항생제코드
    Dim strDrugNm    As String   '수신한 항생제명
    Dim strVolumn    As String   '수신한 항생제 함량
    Dim strRstCd     As String   '수신한 항생제 결과코드
    Dim strTemp      As String
    Dim i            As Long
    
    strRcvBuf = mIntLib.Buffers(1).Buffers
    aryTemp1 = Split(strRcvBuf, GS)
    
    '## Replace후 첫 5자가 msrst이 아니면 Exit
    aryTemp2 = Split(Replace$(aryTemp1(0), RS, ""), FS)
    
    For i = LBound(aryTemp2) To UBound(aryTemp2)
        strTemp = aryTemp2(i)
        strCode = Mid$(strTemp, 1, 2)
        Select Case strCode
            Case "ci"   '## Vitek No
                Set objVitek = New clsIISIntInfo
                objVitek.VitekNo = Format$(Mid$(strTemp, 3), "000000")
                
            Case "o1"   '## 균명(약어)
                objVitek.MnmNm = Mid$(strTemp, 3)
                Call objVitek.GetMnmCd
                
            Case "o2"   '## 균명(전체)
                objVitek.MnmNmFull = Mid$(strTemp, 3)
                
            Case "a1"   '## 항생제코드
                strDrugCd = UCase(Mid$(strTemp, 3))
                
            Case "a2"   '## 항생제명
                strDrugNm = Mid$(strTemp, 3)
                
            Case "a3"   '## 함량
                strVolumn = Mid$(strTemp, 3)
                strVolumn = Replace$(strVolumn, "<=", "≤")
                strVolumn = Replace$(strVolumn, ">=", "≥")
                
            Case "a4"   '## 결과코드
                If strDrugCd = "BETA" Or strDrugCd = "ESBL" Then
                    strRstCd = "  "
                Else
                    strRstCd = UCase$(Mid$(strTemp, 3))
                End If
                
                If strDrugCd <> "OXID" Then
                    Call objVitek.Drugs.Add(strDrugCd, strDrugNm, strVolumn, strRstCd)
                End If
        End Select
    Next i
    
    '## 항생제결과가 있으면 결과저장
    If objVitek.Drugs.Count > 0 Then
        Call SaveServer(objVitek)
    End If
    
    Set objVitek = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 결과판정, 결과저장, 화면표시
'   인수 :
'       - pVitek : Vitek장비에서 수신한 결과저장 클래스
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pVitek As clsIISIntInfo)
    Dim objAccInfo  As clsIISAccInfo    '접수정보 클래스
    Dim vBarNo      As Variant          'Spread의 바코드번호
    Dim strBarno    As String           '바코드번호
    Dim i           As Long
    
    Me.MousePointer = vbHourglass
    
    '## 전송받은 결과를 임시Spread에 저장
    Call SetTemp(pVitek)
    
    Set objAccInfo = mIntLib.GetAccInfoByAccNo(IISMICWA, pVitek.GetAccDt, pVitek.GetAccSeq)
    If objAccInfo Is Nothing Then
        '## 접수정보가 없을때 결과표시
        Call SetComplete1(pVitek)
        Call tblComplete_Click(1, tblComplete.ActiveRow)
    Else
        '## 접수정보가 있을때 결과표시
        Call SetComplete2(objAccInfo, pVitek)
        Call tblComplete_Click(1, tblComplete.ActiveRow)
        strBarno = objAccInfo.GetBarNo
        pVitek.BarNo = strBarno
        Set objAccInfo = Nothing
        
        '## 결과저장
        Call mIntLib.SaveMICResult(pVitek)
        Call mIntLib.RemoveX(strBarno)
        StatusBar.Panels(2).Text = "검체번호:" & strBarno & " 를 정상적으로 결과저장 했습니다."
    End If
    
    '## tblReady에서 전송된 검체삭제
    With tblReady
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
            If CStr(vBarNo) = strBarno Then
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

    If pBarNo = "" Then Exit Sub

    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    
    If Not (objAccInfo Is Nothing) Then
        '## tblReady 정보표시
        Call SetReady(objAccInfo)
        Set objAccInfo = Nothing
    End If
    txtBarNo.Text = "": txtBarNo.SetFocus
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 해당 접수번호에 대한 접수정보 조회, tblReady, tblResult에 표시
'   인수 :
'       - pBarNo : 바코드번호
'-----------------------------------------------------------------------------'
Private Sub GetOrderX(ByVal pWorkarea As String, ByVal pAccDt As String, ByVal pAccSeq As String)
    Dim objAccInfo As clsIISAccInfo     '접수내역 클래스
    Dim vBarNo     As Variant           'Spread의 바코드번호

    Set objAccInfo = mIntLib.GetAccInfoByAccNo(pWorkarea, pAccDt, CLng(pAccSeq))
    If Not (objAccInfo Is Nothing) Then
        '## tblReady 정보표시
        Call SetReady(objAccInfo)
        Set objAccInfo = Nothing
    End If
    Call SetAccNo:  txtAccSeq.Text = ""
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

        .Col = TReadyEnum.ccNo:     .Text = GetVitekNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
        .Col = TReadyEnum.ccBarNo:  .Text = pAccInfo.GetBarNo
        .Col = TReadyEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
        .Col = TReadyEnum.ccPtId:   .Text = pAccInfo.PtId
        .Col = TReadyEnum.ccName:   .Text = pAccInfo.Name
        .Col = TReadyEnum.ccTestNm: .Text = pAccInfo.MICResult.TestNm
        .Col = TReadyEnum.ccSpcNm:  .Text = pAccInfo.SpcNm
        .Col = TReadyEnum.ccSex:    .Text = pAccInfo.Sex
        .Col = TReadyEnum.ccAge:    .Text = pAccInfo.Age
        .Col = TReadyEnum.ccDoctNm: .Text = pAccInfo.OrdDoctNm
        .Col = TReadyEnum.ccDeptNm: .Text = pAccInfo.DeptNm
        .Col = TReadyEnum.ccWardNm: .Text = pAccInfo.WardNm
        .Col = TReadyEnum.ccWorkSheet
        .Text = pAccInfo.MICResult.WSBody.WsCd & "-" & pAccInfo.MICResult.WSBody.WsUnit
        
        Call .SetActiveCell(1, .DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblTemp에 정보표시
'   인수 :
'       - pVitek : Vitek장비에서 수신한 결과저장 클래스
'-----------------------------------------------------------------------------'
Private Sub SetTemp(ByVal pVitek As clsIISIntInfo)
    Dim objDrug As clsIISMICDrug   '항생제결과 클래스
    Dim i       As Long
    
    With tblTemp
        .Row = .DataRowCnt + 1
        .Col = TTempEnum.ccNo:      .Text = pVitek.VitekNo
        .Col = TTempEnum.ccMnmCd:   .Text = IIf(pVitek.MnmCd = "", pVitek.MnmNm, pVitek.MnmCd)
        .Col = TTempEnum.ccMnmNm:   .Text = pVitek.MnmNmFull
        .Col = TTempEnum.ccCount:   .Text = pVitek.Drugs.Count
        
        For Each objDrug In pVitek.Drugs
            If .MaxCols <= .DataColCnt Then
                .MaxCols = .MaxCols + 1
            End If
            .Col = TTempEnum.ccResult + i
            .Text = objDrug.DrugCd & DIV & objDrug.DrugNm & DIV & objDrug.RstCd & DIV & _
                    objDrug.Volumn
            i = i + 1
        Next
        Set objDrug = Nothing
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblComplete에 정보표시 (접수정보가 없을때)
'   인수 :
'       - pVitek : Vitek장비에서 수신한 결과저장 클래스
'-----------------------------------------------------------------------------'
Private Sub SetComplete1(ByVal pVitek As clsIISIntInfo)
    Dim vVitekNo As Variant     'VitekNo
    Dim i        As Long
    
    With tblComplete
        '## 검사완료리스트에 같은 VitekNo가 있는지 조회
        For i = 1 To .DataRowCnt
            Call .GetText(TCompleteEnum.ccNo, i, vVitekNo)
            
            If pVitek.VitekNo = CStr(vVitekNo) Then
                Call .SetActiveCell(1, i)
                Exit Sub
            End If
        Next i
        
        '## VitekNo가 없으면 추가
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If
        .Col = TCompleteEnum.ccNo:  .Text = pVitek.VitekNo
        
        Call .SetActiveCell(1, .DataRowCnt)
        lblCompleteCnt.Caption = CStr(.DataRowCnt)
    End With
End Sub

'-----------------------------------------------------------------------------'
'   기능 : tblComplete에 정보표시 (접수정보가 있을때)
'   인수 :
'       - pAccInfo : 접수내역 클래스
'       - pVitek : Vitek장비에서 수신한 결과저장 클래스
'-----------------------------------------------------------------------------'
Private Sub SetComplete2(ByVal pAccInfo As clsIISAccInfo, ByVal pVitek As clsIISIntInfo)
    Dim vVitekNo As Variant     'VitekNo
    Dim i        As Long
    
    With tblComplete
        '## 검사완료리스트에 같은 VitekNo가 있는지 조회
        For i = 1 To .DataRowCnt
            Call .GetText(TCompleteEnum.ccNo, i, vVitekNo)
            
            If pVitek.VitekNo = CStr(vVitekNo) Then
                Call .SetActiveCell(1, i)
                Exit Sub
            End If
        Next i
        
        '## VitekNo가 없으면 추가
        If .MaxRows <= .DataRowCnt Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
        Else
            .Row = .DataRowCnt + 1
        End If
        .Col = TCompleteEnum.ccNo:     .Text = pVitek.VitekNo
        .Col = TCompleteEnum.ccBarNo:  .Text = pAccInfo.GetBarNo
        .Col = TCompleteEnum.ccAccNo:  .Text = mGetAccNo(pAccInfo.Workarea, pAccInfo.AccDt, pAccInfo.AccSeq)
        .Col = TCompleteEnum.ccPtId:   .Text = pAccInfo.PtId
        .Col = TCompleteEnum.ccName:   .Text = pAccInfo.Name
        .Col = TCompleteEnum.ccTestNm: .Text = pAccInfo.MICResult.TestNm
        .Col = TCompleteEnum.ccSpcNm:  .Text = pAccInfo.SpcNm
        .Col = TCompleteEnum.ccSex:    .Text = pAccInfo.Sex
        .Col = TCompleteEnum.ccAge:    .Text = pAccInfo.Age
        .Col = TCompleteEnum.ccDoctNm: .Text = pAccInfo.OrdDoctNm
        .Col = TCompleteEnum.ccDeptNm: .Text = pAccInfo.DeptNm
        .Col = TCompleteEnum.ccWardNm: .Text = pAccInfo.WardNm
        .Col = TCompleteEnum.ccWorkSheet
        .Text = pAccInfo.MICResult.WSBody.WsCd & "-" & pAccInfo.MICResult.WSBody.WsUnit
        
        Call .SetActiveCell(1, .DataRowCnt)
        lblCompleteCnt.Caption = CStr(.DataRowCnt)
    End With
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
            Error.SetLog App.EXEName, "frmIISVitekII", "GetEqpComm", strErrMsg, Now
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
        Error.SetLog App.EXEName, "frmIISVitekII", "GetEqpComm", strErrMsg, Now
        Call mIntLib_EqpError("E005")
    End If
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 컨트롤 초기화
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtBarNo.Text = "":          Call mTblClear(tblReady)
    Call mTblClear(tblComplete): Call mTblClear(tblMnm)
    Call mTblClear(tblSensi):    Call mTblClear(tblTemp)
    lblReadyCnt.Caption = "0":   lblCompleteCnt.Caption = "0"
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
                    Call mTblClear(tblMnm)
                    Call mTblClear(tblSensi)
                End If
                lblReadyCnt.Caption = CStr(tblReady.DataRowCnt)
            End With
        Case DELETEALL  '## Delete All
            Call mIntLib.AccInfos.RemoveAll
            Call mTblClear(tblReady)
            lblReadyCnt.Caption = "0"
    End Select
End Sub

'------------------------------------------------------------------'
'   기능 : 접수번호를 이용해 Vitek No를 조회
'   인수 :
'       - pWorkarea : Workarea
'       - pAccDt    : 접수일자
'       - pAccSeq   : 접수순번
'   반환 : Vitek No
'------------------------------------------------------------------'
Private Function GetVitekNo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    '## 07-200409-7001 --> 097001
    GetVitekNo = Mid$(pAccDt, 5, 2) & Format$(CStr(pAccSeq), "0000")
End Function

'------------------------------------------------------------------'
'   기능 : 접수번의 Workarea, AccDt 부분의 자동표시
'------------------------------------------------------------------'
Private Sub SetAccNo()
    txtWorkarea.Text = IISMICWA
    txtAccDt.Text = Format$(Now, "YYMM")
    txtAccSeq.Text = ""
End Sub

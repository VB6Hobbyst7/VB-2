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
      Name            =   "����"
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
   StartUpPosition =   3  'Windows �⺻��
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
         Name            =   "����"
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
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
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
      Alignment       =   2  '��� ����
      BorderStyle     =   0  '����
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
      Alignment       =   2  '��� ����
      BorderStyle     =   0  '����
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
      Alignment       =   2  '��� ����
      BorderStyle     =   0  '����
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
      Caption         =   "������ȣ"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10290
      Style           =   1  '�׷���
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� ������"
      Appearance      =   0
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00DBE6E6&
      Caption         =   "ȭ������(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12698
      Style           =   1  '�׷���
      TabIndex        =   7
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdAlarm 
      BackColor       =   &H00DBE6E6&
      Caption         =   "&Alarm"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11483
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   8567
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13913
      Style           =   1  '�׷���
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
         Name            =   "����"
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� �˻��� ����Ʈ"
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� �˻�Ϸ� ����Ʈ"
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
         Name            =   "����"
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
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�� �׻��� ����"
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
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "���� :"
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   0  '����
      Caption         =   "���� :"
      BeginProperty Font 
         Name            =   "����"
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
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00000000&
      FillColor       =   &H80000005&
      Height          =   315
      Left            =   1440
      Top             =   510
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��� ����
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
      Alignment       =   2  '��� ����
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
      BackStyle       =   0  '����
      Caption         =   "���ڵ��ȣ : "
      BeginProperty Font 
         Name            =   "����"
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
'   ���ϸ�  : frmIISVitekII.frm
'   �ۼ���  : �̻��
'   ��  ��  : Vitek II �����
'   �ۼ���  : 2005-01-31
'   ��  ��  :
'       1. 1.0.1: �̻��(2005-02-25)
'   ��  ��  :
'       1. ��������
'-----------------------------------------------------------------------------'

Option Explicit

'## tblReady�� Column Enum
Private Enum TReadyEnum
    ccOrdChk = 1:       ccNo = 2
    ccBarNo = 3:        ccAccNo = 4
    ccPtId = 5:         ccName = 6
    ccTestNm = 7:       ccSpcNm = 8
    ccSex = 9:          ccAge = 10
    ccDoctNm = 11:      ccDeptNm = 12
    ccWardNm = 13:      ccWorkSheet = 14
End Enum

'## tblComplete�� Column Enum
Private Enum TCompleteEnum
    ccNo = 1:           ccBarNo = 2
    ccAccNo = 3:        ccPtId = 4
    ccName = 5:         ccTestNm = 6
    ccSpcNm = 7:        ccSex = 8
    ccAge = 9:          ccDoctNm = 10
    ccDeptNm = 11:      ccWardNm = 12
    ccWorkSheet = 13
End Enum

'## tblTemp�� Column Enum
Private Enum TTempEnum
    ccNo = 1
    ccMnmCd = 2
    ccMnmNm = 3
    ccCount = 4
    ccResult = 5
End Enum

''## tblMnm�� Column Enum
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

'## Datalog Field ���
Private Const RS As String = ""    'Record Separator
Private Const GS As String = ""    'Group Separator
Private Const FS As String = "|"    'Field Separator

'## Popup Menu ID
Private Const DELETE    As Long = 1
Private Const DELETEALL As Long = 2

Private WithEvents mIntLib  As clsIISInterface   '�������̽� Ŭ����
Attribute mIntLib.VB_VarHelpID = -1
Private WithEvents mPopup   As clsIISPopup       '�˾��޴�
Attribute mPopup.VB_VarHelpID = -1

Private mIntErrors  As clsIISIntErrors           '�������̽� ���� �÷���
Private mOrder      As clsIISIntOrder            '�������� Ŭ����

Private mEqpCd  As String   '����ڵ�
Private mEqpKey As String   '���Ű

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
    Dim objErrorShow As clsIISErrorShow     '������ ǥ�� Ŭ����

    Set objErrorShow = New clsIISErrorShow
    Call objErrorShow.ShowErrors(mIntErrors)
    Set objErrorShow = Nothing

    '## ������ ������ ��ư���� �������, ������ ��� ������
    cmdAlarm.BackColor = IIf(mIntErrors.Count = 0, &HF4F0F2, vbRed)
    
    '## 1.0.1: �̻��(2005-02-25)
    '   - Alarmâ�� ������ ��Ŀ���� txtBarNo Or txtAccSeq�� �̵�
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
    Dim vOrdChk As Variant  'Spread�� �������ۿ���
    Dim i       As Long
    
    '## ��Ʈ�� ���µǾ� ���� ������ ����ǥ��
    If MSComm.PortOpen = False Then
        MsgBox "��Ʈ�� �������� �ʽ��ϴ�.", vbCritical, "����"
        Exit Sub
    End If

    With tblReady
        If .DataRowCnt < 1 Then Exit Sub
        
        '## �۽��� ��ü���� �ľ�!
        mOrder.SendCnt = 0
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccOrdChk, i, vOrdChk)
            
            If CStr(vOrdChk) = "" Then
                mOrder.SendCnt = mOrder.SendCnt + 1
            End If
        Next i
    End With
        
    '## ENQ ����
    MSComm.Output = ENQ
    Call mIntLib.WriteLog(ENQ, ccPCLog)
    mIntLib.State = "Q"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub chkAccNo_Click()
    If chkAccNo.Value = 1 Then
        lblNo.Caption = "������ȣ : "
        txtBarNo.Visible = False
        Shape1.Visible = True:      txtWorkarea.Visible = True:
        txtAccDt.Visible = True:    txtAccSeq.Visible = True:
        Label1.Visible = True:      Label1.ZOrder 0
        Label2.Visible = True:      Label2.ZOrder 0
        Call SetAccNo:              txtAccSeq.SetFocus
    Else
        lblNo.Caption = "���ڵ��ȣ : "
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
    '## �ش� ���ڵ��ȣ�� ���� �������� ��ȸ
    If KeyCode = vbKeyReturn Then
        Me.MousePointer = vbHourglass
        
        Call GetOrder(Trim(txtBarNo.Text))
        lblReadyCnt.Caption = CStr(tblReady.DataRowCnt)
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub txtBarNo_KeyPress(KeyAscii As Integer)
    '## ���ڸ� �Է�
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
        
        '## 2100�� ���ʹ� �����α׷� ������!
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
    Dim vBarNo      As Variant          'Spread�� ���ڵ��ȣ
    Dim strInfo     As String           '��������
    
    If Row = 0 Then Exit Sub
    With tblReady
        Call .GetText(TReadyEnum.ccBarNo, Row, vBarNo)
        If Trim$(CStr(vBarNo)) = "" Then Exit Sub
        
        .Row = Row
        strInfo = vbCrLf
        .Col = TReadyEnum.ccTestNm
        strInfo = strInfo & Space(2) & "�� �� �� : " & .Text & vbCrLf
        .Col = TReadyEnum.ccSpcNm
        strInfo = strInfo & Space(2) & "�� ü �� : " & .Text & vbCrLf
        .Col = TReadyEnum.ccSex
        strInfo = strInfo & Space(2) & "��    �� : " & .Text & vbCrLf
        .Col = TReadyEnum.ccAge
        strInfo = strInfo & Space(2) & "��    �� : " & .Text & vbCrLf
        .Col = TReadyEnum.ccDoctNm
        strInfo = strInfo & Space(2) & "ó �� �� : " & .Text & vbCrLf
        .Col = TReadyEnum.ccDeptNm
        strInfo = strInfo & Space(2) & "�� �� �� : " & .Text & vbCrLf
        .Col = TReadyEnum.ccWardNm
        strInfo = strInfo & Space(2) & "��    �� : " & .Text & vbCrLf
        .Col = TReadyEnum.ccWorkSheet
        strInfo = strInfo & Space(2) & "WorkSheet Unit : " & .Text & vbCrLf
        
        MultiLine = 1
        TipWidth = 6000
        TipText = strInfo
        Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
End Sub

Private Sub tblComplete_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vVitekNo1   As Variant      'Spread�� Vitek No1(tblComplete)
    Dim vVitekNo2   As Variant      'Spread�� Vitek No2(tblTemp)
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
        
        '## �ش� ���ڵ��� �׻������ ��ȸ
        If .DataRowCnt > 0 Then
            Call .SetActiveCell(1, 1)
            Call tblMnm_Click(1, 1)
        End If
    End With
End Sub

Private Sub tblComplete_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    Dim vBarNo      As Variant          'Spread�� ���ڵ��ȣ
    Dim strInfo     As String           '��������
    
    If Row = 0 Then Exit Sub
    With tblComplete
        Call .GetText(TReadyEnum.ccBarNo, Row, vBarNo)
        If Trim$(CStr(vBarNo)) = "" Then Exit Sub
        
        .Row = Row
        strInfo = vbCrLf
        .Col = TCompleteEnum.ccTestNm
        strInfo = strInfo & Space(2) & "�� �� �� : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccSpcNm
        strInfo = strInfo & Space(2) & "�� ü �� : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccSex
        strInfo = strInfo & Space(2) & "��    �� : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccAge
        strInfo = strInfo & Space(2) & "��    �� : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccDoctNm
        strInfo = strInfo & Space(2) & "ó �� �� : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccDeptNm
        strInfo = strInfo & Space(2) & "�� �� �� : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccWardNm
        strInfo = strInfo & Space(2) & "��    �� : " & .Text & vbCrLf
        .Col = TCompleteEnum.ccWorkSheet
        strInfo = strInfo & Space(2) & "WorkSheet Unit : " & .Text & vbCrLf
        
        MultiLine = 1
        TipWidth = 6000
        TipText = strInfo
        Call .SetTextTipAppearance("����ü", 9, False, False, &HEEFDF2, &H996666)
        ShowTip = True
    End With
End Sub

Private Sub tblMnm_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vVitekNo1   As Variant  'Spread�� Vitek No1(tblMnm)
    Dim vVitekNo2   As Variant  'Spread�� Vitek No2(tblTemp)
    Dim vMnmCd1     As Variant  'Spread�� ���ڵ�1(tblMnm)
    Dim vMnmCd2     As Variant  'Spread�� ���ڵ�2(tblTemp)
    Dim vCount      As Variant  'Spread�� �׻�������(tblTemp)
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
                    Case 1      '## ENQ, ACK ���
                        Select Case BufChar
                            Case ENQ
                                mIntLib.BufCnt = 1
                                Call mIntLib.ClearBuffer
                                
                                MSComm.Output = ACK
                                Call mIntLib.WriteLog(ACK, ccPCLog)
                                mIntLib.Phase = 2
                            Case ACK
                                If mIntLib.State = "Q" Then     '## ENQ ������
                                    Call SendOrder
                                ElseIf mIntLib.State = "C" Then '## CheckSum ������
                                    '## ������ ��ü Check ǥ��
                                    Call tblReady.SetText(TReadyEnum.ccOrdChk, mOrder.Seq, "��")
                                    mOrder.SendCnt = mOrder.SendCnt - 1
                                    
                                    '## ETX ����
                                    MSComm.Output = ETX
                                    Call mIntLib.WriteLog(ETX, ccPCLog)
                                    
                                    '## EOT ����
                                    MSComm.Output = EOT
                                    Call mIntLib.WriteLog(EOT, ccPCLog)
                                    
                                    '## ������ ��ü�� ������ ENQ����
                                    If mOrder.SendCnt > 0 Then
                                        Call mSleep(1000)
                                        mIntLib.State = "Q"
                                        MSComm.Output = ENQ
                                        Call mIntLib.WriteLog(ENQ, ccPCLog)
                                    End If
                                End If
                        End Select
                    Case 2      '## GS ���
                        Select Case BufChar
                            Case STX
                            Case GS
                                mIntLib.Phase = 3
                            Case Else
                                Call mIntLib.AddBuffer(BufChar)
                        End Select
                    Case 3      '## CheckSum ���
                        lngCheckSum = lngCheckSum + 1
                        If lngCheckSum = 2 Then
                            MSComm.Output = ACK
                            Call mIntLib.WriteLog(ACK, ccPCLog)
                            mIntLib.Phase = 4
                        End If
                    Case 4      '## CheckSum ���
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
            EVMsg$ = "CTS ���� ����"
        Case comEvDSR
            EVMsg$ = "DSR ���� ����"
        Case comEvCD
            EVMsg$ = "CD ���� ����"
        Case comEvRing
            EVMsg$ = "��ȭ ���� �︮�� ��"
        Case comEvEOF
            EVMsg$ = "EOF ����"

        '���� �޽���
        Case comBreak
            ERMsg$ = "�ߴ� ��ȣ ����"
        Case comCDTO
            ERMsg$ = "�ݼ��� ���� �ð� �ʰ�"
        Case comCTSTO
            ERMsg$ = "CTS �ð� �ʰ�"
        Case comDCB
            ERMsg$ = "DCB �˻� ����"
        Case comDSRTO
            ERMsg$ = "DSR �ð� �ʰ�"
        Case comFrame
            ERMsg$ = "�����̹� ����"
        Case comOverrun
            ERMsg$ = "�и�Ƽ ����"
        Case comRxOver
            ERMsg$ = "���� ���� �ʰ�"
        Case comRxParity
            ERMsg$ = "�и�Ƽ ����"
        Case comTxFull
            ERMsg$ = "���� ���ۿ� ������ ����"
        Case Else
            ERMsg$ = "�� �� ���� ���� �Ǵ� �̺�Ʈ"
    End Select

    If Len(EVMsg$) Then
        StatusBar.Panels(2).Text = EVMsg$
    ElseIf Len(ERMsg$) Then
        StatusBar.Panels(2).Text = ERMsg$
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
    Dim vOrdChk     As Variant          'Spread�� ������������
    Dim vBarNo      As Variant          'Spread�� ���ڵ��ȣ
    Dim vVitekNo    As Variant          'Spread�� VitekNo
    Dim strOutput   As String           '�۽��� ������
    Dim blnSend     As Boolean          '�������� ����
    Dim i           As Long
    
    With tblReady
        '## Spread���� ������ �������� ���� ��ü�� �˻��� ��������
        For i = 1 To .DataRowCnt
            Call .GetText(TReadyEnum.ccOrdChk, i, vOrdChk)
            
            If CStr(vOrdChk) = "" Then
                Call .GetText(TReadyEnum.ccNo, i, vVitekNo)
                Call .GetText(TReadyEnum.ccBarNo, i, vBarNo)
                
                Set objAccInfo = mIntLib.GetAccInfoX(CStr(vBarNo))
                If Not (objAccInfo Is Nothing) Then
                    '## �������� Ŭ���� �ʱ�ȭ
                    blnSend = True
                    mOrder.ClsClear
                    mOrder.Seq = i
                    
                    '## 1.STX ����
                    strOutput = STX & vbCrLf
                    MSComm.Output = strOutput
                    Call mIntLib.WriteLog(strOutput, ccPCLog)
                    
                    '## 2.�������ڿ� ����
                    strOutput = mOrder.GetOrder(objAccInfo, CStr(vVitekNo))
                    MSComm.Output = strOutput
                    Call mIntLib.WriteLog(strOutput, ccPCLog)
                    
                    '## 3.CheckSum ����
                    strOutput = GS & mOrder.CheckSum & vbCrLf
                    MSComm.Output = strOutput
                    Call mIntLib.WriteLog(strOutput, ccPCLog)
                    
                    mIntLib.State = "C"
                    Set objAccInfo = Nothing
                    Exit Sub
                End If
            End If
        Next i
        
        '## �߸��� ������ ��� EOT �����Ͽ� �������
        If blnSend = False Then
            mIntLib.State = ""
            MSComm.Output = EOT
            Call mIntLib.WriteLog(EOT, ccPCLog)
        End If
    End With
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���κ��� ������ ������ ����
'-----------------------------------------------------------------------------'
Private Sub EditRcvData()
    Dim objVitek     As clsIISIntInfo   '��񿡼� ������ ������� Ŭ����
    
    Dim aryTemp1()   As String
    Dim aryTemp2()   As String
    Dim strRcvBuf    As String   '������ Data
    Dim strCode      As String   '������ Field Code
    Dim strDrugCd    As String   '������ �׻����ڵ�
    Dim strDrugNm    As String   '������ �׻�����
    Dim strVolumn    As String   '������ �׻��� �Է�
    Dim strRstCd     As String   '������ �׻��� ����ڵ�
    Dim strTemp      As String
    Dim i            As Long
    
    strRcvBuf = mIntLib.Buffers(1).Buffers
    aryTemp1 = Split(strRcvBuf, GS)
    
    '## Replace�� ù 5�ڰ� msrst�� �ƴϸ� Exit
    aryTemp2 = Split(Replace$(aryTemp1(0), RS, ""), FS)
    
    For i = LBound(aryTemp2) To UBound(aryTemp2)
        strTemp = aryTemp2(i)
        strCode = Mid$(strTemp, 1, 2)
        Select Case strCode
            Case "ci"   '## Vitek No
                Set objVitek = New clsIISIntInfo
                objVitek.VitekNo = Format$(Mid$(strTemp, 3), "000000")
                
            Case "o1"   '## �ո�(���)
                objVitek.MnmNm = Mid$(strTemp, 3)
                Call objVitek.GetMnmCd
                
            Case "o2"   '## �ո�(��ü)
                objVitek.MnmNmFull = Mid$(strTemp, 3)
                
            Case "a1"   '## �׻����ڵ�
                strDrugCd = UCase(Mid$(strTemp, 3))
                
            Case "a2"   '## �׻�����
                strDrugNm = Mid$(strTemp, 3)
                
            Case "a3"   '## �Է�
                strVolumn = Mid$(strTemp, 3)
                strVolumn = Replace$(strVolumn, "<=", "��")
                strVolumn = Replace$(strVolumn, ">=", "��")
                
            Case "a4"   '## ����ڵ�
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
    
    '## �׻�������� ������ �������
    If objVitek.Drugs.Count > 0 Then
        Call SaveServer(objVitek)
    End If
    
    Set objVitek = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �������, �������, ȭ��ǥ��
'   �μ� :
'       - pVitek : Vitek��񿡼� ������ ������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SaveServer(ByVal pVitek As clsIISIntInfo)
    Dim objAccInfo  As clsIISAccInfo    '�������� Ŭ����
    Dim vBarNo      As Variant          'Spread�� ���ڵ��ȣ
    Dim strBarno    As String           '���ڵ��ȣ
    Dim i           As Long
    
    Me.MousePointer = vbHourglass
    
    '## ���۹��� ����� �ӽ�Spread�� ����
    Call SetTemp(pVitek)
    
    Set objAccInfo = mIntLib.GetAccInfoByAccNo(IISMICWA, pVitek.GetAccDt, pVitek.GetAccSeq)
    If objAccInfo Is Nothing Then
        '## ���������� ������ ���ǥ��
        Call SetComplete1(pVitek)
        Call tblComplete_Click(1, tblComplete.ActiveRow)
    Else
        '## ���������� ������ ���ǥ��
        Call SetComplete2(objAccInfo, pVitek)
        Call tblComplete_Click(1, tblComplete.ActiveRow)
        strBarno = objAccInfo.GetBarNo
        pVitek.BarNo = strBarno
        Set objAccInfo = Nothing
        
        '## �������
        Call mIntLib.SaveMICResult(pVitek)
        Call mIntLib.RemoveX(strBarno)
        StatusBar.Panels(2).Text = "��ü��ȣ:" & strBarno & " �� ���������� ������� �߽��ϴ�."
    End If
    
    '## tblReady���� ���۵� ��ü����
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
'   ��� : �ش� ���ڵ��ȣ�� ���� �������� ��ȸ, tblReady, tblResult�� ǥ��
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarNo As String)
    Dim objAccInfo As clsIISAccInfo     '�������� Ŭ����

    If pBarNo = "" Then Exit Sub

    Set objAccInfo = mIntLib.GetAccInfo(pBarNo)
    
    If Not (objAccInfo Is Nothing) Then
        '## tblReady ����ǥ��
        Call SetReady(objAccInfo)
        Set objAccInfo = Nothing
    End If
    txtBarNo.Text = "": txtBarNo.SetFocus
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ������ȣ�� ���� �������� ��ȸ, tblReady, tblResult�� ǥ��
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'-----------------------------------------------------------------------------'
Private Sub GetOrderX(ByVal pWorkarea As String, ByVal pAccDt As String, ByVal pAccSeq As String)
    Dim objAccInfo As clsIISAccInfo     '�������� Ŭ����
    Dim vBarNo     As Variant           'Spread�� ���ڵ��ȣ

    Set objAccInfo = mIntLib.GetAccInfoByAccNo(pWorkarea, pAccDt, CLng(pAccSeq))
    If Not (objAccInfo Is Nothing) Then
        '## tblReady ����ǥ��
        Call SetReady(objAccInfo)
        Set objAccInfo = Nothing
    End If
    Call SetAccNo:  txtAccSeq.Text = ""
End Sub

'-----------------------------------------------------------------------------'
'   ��� : tblReady�� ����ǥ��
'   �μ� :
'       - pAccInfo : �������� Ŭ����
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
'   ��� : tblTemp�� ����ǥ��
'   �μ� :
'       - pVitek : Vitek��񿡼� ������ ������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetTemp(ByVal pVitek As clsIISIntInfo)
    Dim objDrug As clsIISMICDrug   '�׻������ Ŭ����
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
'   ��� : tblComplete�� ����ǥ�� (���������� ������)
'   �μ� :
'       - pVitek : Vitek��񿡼� ������ ������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetComplete1(ByVal pVitek As clsIISIntInfo)
    Dim vVitekNo As Variant     'VitekNo
    Dim i        As Long
    
    With tblComplete
        '## �˻�ϷḮ��Ʈ�� ���� VitekNo�� �ִ��� ��ȸ
        For i = 1 To .DataRowCnt
            Call .GetText(TCompleteEnum.ccNo, i, vVitekNo)
            
            If pVitek.VitekNo = CStr(vVitekNo) Then
                Call .SetActiveCell(1, i)
                Exit Sub
            End If
        Next i
        
        '## VitekNo�� ������ �߰�
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
'   ��� : tblComplete�� ����ǥ�� (���������� ������)
'   �μ� :
'       - pAccInfo : �������� Ŭ����
'       - pVitek : Vitek��񿡼� ������ ������� Ŭ����
'-----------------------------------------------------------------------------'
Private Sub SetComplete2(ByVal pAccInfo As clsIISAccInfo, ByVal pVitek As clsIISIntInfo)
    Dim vVitekNo As Variant     'VitekNo
    Dim i        As Long
    
    With tblComplete
        '## �˻�ϷḮ��Ʈ�� ���� VitekNo�� �ִ��� ��ȸ
        For i = 1 To .DataRowCnt
            Call .GetText(TCompleteEnum.ccNo, i, vVitekNo)
            
            If pVitek.VitekNo = CStr(vVitekNo) Then
                Call .SetActiveCell(1, i)
                Exit Sub
            End If
        Next i
        
        '## VitekNo�� ������ �߰�
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
'   ��� : ����� ������ȸ, ��Ʈ Open
'-----------------------------------------------------------------------------'
Private Sub GetEqpComm()
    Dim objComm     As clsIISEqpComm    '��ż��� Ŭ����
    Dim strErrMsg   As String           '�����޽���

    '## ��ż��� ������ȸ
    Set objComm = mIntLib.GetEqpComm
    If objComm Is Nothing Then Exit Sub

    With objComm
        MSComm.CommPort = .Port
        MSComm.Settings = .GetSettings
    End With
    Set objComm = Nothing

On Error GoTo Errors
    '## ��Ʈ Open
    With MSComm
        '## �̹� ��Ʈ�� �������
        If .PortOpen Then
            strErrMsg = mEqpCd & " ����� �����Ʈ�� �̹� �����ֽ��ϴ�."
            Error.SetLog App.EXEName, "frmIISVitekII", "GetEqpComm", strErrMsg, Now
            Call mIntLib_EqpError("E004")
            Exit Sub
        End If

        .RThreshold = 1
        .SThreshold = 1
        .RTSEnable = True
        .PortOpen = True
    End With

    '## �������� ���������� ����
    Call mIntLib.DelHistoryData
    Exit Sub

Errors:
    '## �ٸ� ��ġ���� ��Ʈ�� ����ϴ� ���
    If Err.Number = 8005 Then
        strErrMsg = mEqpCd & " ��� ������ ��Ʈ�� �̹� ������Դϴ�."
        Error.SetLog App.EXEName, "frmIISVitekII", "GetEqpComm", strErrMsg, Now
        Call mIntLib_EqpError("E005")
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��Ʈ�� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    txtBarNo.Text = "":          Call mTblClear(tblReady)
    Call mTblClear(tblComplete): Call mTblClear(tblMnm)
    Call mTblClear(tblSensi):    Call mTblClear(tblTemp)
    lblReadyCnt.Caption = "0":   lblCompleteCnt.Caption = "0"
End Sub

'------------------------------------------------------------------'
'   ��� : ����� ���� ����ó��
'------------------------------------------------------------------'
Private Sub mIntLib_EqpError(ByVal pCode As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey)
End Sub

'------------------------------------------------------------------'
'   ��� : ��ü���� ����ó��1
'------------------------------------------------------------------'
Private Sub mIntLib_SpcError(ByVal pCode As String, ByVal pBarNo As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo)
End Sub

'------------------------------------------------------------------'
'   ��� : ��ü���� ����ó��2
'------------------------------------------------------------------'
Private Sub mIntLib_SpcErrorX(ByVal pCode As String, ByVal pBarNo As String, ByVal pPtId As String, ByVal pName As String)
    cmdAlarm.BackColor = vbRed
    Call mIntErrors.Add(pCode, mEqpCd, mEqpKey, pBarNo, pPtId, pName)
End Sub

'------------------------------------------------------------------'
'   ��� : Popup �޴� Click �̺�Ʈ
'------------------------------------------------------------------'
Private Sub mPopup_Click(ByVal vMenuID As Long)
    Dim vBarNo      As Variant  'Spread�� ���ڵ��ȣ
    Dim strSpcYy    As String   '��ü����
    Dim lngSpcNo    As Long     '��ü��ȣ

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
'   ��� : ������ȣ�� �̿��� Vitek No�� ��ȸ
'   �μ� :
'       - pWorkarea : Workarea
'       - pAccDt    : ��������
'       - pAccSeq   : ��������
'   ��ȯ : Vitek No
'------------------------------------------------------------------'
Private Function GetVitekNo(ByVal pWorkarea As String, ByVal pAccDt As String, _
                    ByVal pAccSeq As Long) As String
    '## 07-200409-7001 --> 097001
    GetVitekNo = Mid$(pAccDt, 5, 2) & Format$(CStr(pAccSeq), "0000")
End Function

'------------------------------------------------------------------'
'   ��� : �������� Workarea, AccDt �κ��� �ڵ�ǥ��
'------------------------------------------------------------------'
Private Sub SetAccNo()
    txtWorkarea.Text = IISMICWA
    txtAccDt.Text = Format$(Now, "YYMM")
    txtAccSeq.Text = ""
End Sub

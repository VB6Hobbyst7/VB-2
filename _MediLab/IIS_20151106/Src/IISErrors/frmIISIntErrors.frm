VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmIISIntErrors 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "��������"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '������ ���
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&D)"
      Height          =   495
      Left            =   7260
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   7470
      Width           =   1215
   End
   Begin VB.TextBox txtDetail 
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   68
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4515
      Width           =   10845
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      Height          =   495
      Left            =   9690
      Style           =   1  '�׷���
      TabIndex        =   2
      Top             =   7470
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteAll 
      BackColor       =   &H00DBE6E6&
      Caption         =   "��λ���(&A)"
      Height          =   495
      Left            =   8475
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   7470
      Width           =   1215
   End
   Begin FPSpread.vaSpread tblErrors 
      Height          =   4425
      Left            =   75
      TabIndex        =   3
      Top             =   60
      Width           =   10845
      _Version        =   393216
      _ExtentX        =   19129
      _ExtentY        =   7805
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      MaxRows         =   14
      OperationMode   =   2
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   16773087
      SpreadDesigner  =   "frmIISIntErrors.frx":0000
   End
End
Attribute VB_Name = "frmIISIntErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIISIntErrors.frm
'   �ۼ���  : ������
'   ��  ��  : �������� ǥ����
'   �ۼ���  : 2015-10-30
'   ��  ��  : 1.0.0
'-----------------------------------------------------------------------------'

Option Explicit

'## tblErrors�� Column Enum
Private Enum TErrorsEnum
    ccCode = 1
    ccDate = 2
    ccTitle = 3
    ccSeq = 4
End Enum

Private mIntErrors As clsIISIntErrors       '�������̽� ���� �÷���

Public Property Let IntErrors(ByRef vData As clsIISIntErrors)
    Set mIntErrors = vData
End Property

Private Sub Form_Load()
    Me.MousePointer = vbHourglass
    
    Call GetErrors
    Call tblErrors_Click(1, 1)
    
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mIntErrors = Nothing
    Set frmIISIntErrors = Nothing
End Sub

Private Sub cmdDelete_Click()
    Dim vSeq As Variant     'Spread�� Seq
    
    With tblErrors
        Call .GetText(TErrorsEnum.ccSeq, .ActiveRow, vSeq)
        If vSeq <> "" Then
            Call mIntErrors.Remove(CLng(vSeq))
        End If
        
        Call .SetActiveCell(TErrorsEnum.ccSeq, 1)
    End With
    
    Call GetErrors
    Call tblErrors_Click(TErrorsEnum.ccSeq, 1)
End Sub

Private Sub cmdDeleteAll_Click()
    Call mIntErrors.RemoveAll
    Call CtlClear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub tblErrors_Click(ByVal Col As Long, ByVal Row As Long)
    Dim vSeq As Variant     'Spread�� Seq
    
    txtDetail.Text = ""
    
    Call tblErrors.GetText(TErrorsEnum.ccSeq, Row, vSeq)
    If vSeq = "" Then Exit Sub
    
    If mIntErrors.Exist(CLng(vSeq)) Then
        txtDetail.Text = mIntErrors(CLng(vSeq)).GetDescription
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �����÷����� ������ tblErrors�� ǥ��
'-----------------------------------------------------------------------------'
Private Sub GetErrors()
    Dim objIntError As clsIISIntError   '�������̽� ���� Ŭ����
    
    Call CtlClear
    For Each objIntError In mIntErrors
        With tblErrors
            If .DataRowCnt >= .MaxRows Then
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
            Else
                .Row = .DataRowCnt + 1
            End If

            .Col = TErrorsEnum.ccCode:   .Text = objIntError.Code
            .Col = TErrorsEnum.ccDate:   .Text = objIntError.ErrDt
            .Col = TErrorsEnum.ccTitle:  .Text = objIntError.GetTitle
            .Col = TErrorsEnum.ccSeq:    .Text = CStr(objIntError.Seq)
        End With
    Next
    Set objIntError = Nothing
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ��Ʈ�� �ʱ�ȭ
'-----------------------------------------------------------------------------'
Private Sub CtlClear()
    Call tblErrors.ClearRange(1, 1, tblErrors.MaxCols, tblErrors.MaxRows, True)
    txtDetail.Text = ""
End Sub

VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmIISConfig 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '���� ����
   Caption         =   "BEP2000 ����"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5355
   StartUpPosition =   1  '������ ���
   Begin VB.TextBox txtResult 
      BackColor       =   &H00F7FFF7&
      Height          =   330
      Left            =   116
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3340
      Width           =   4725
   End
   Begin VB.TextBox txtOrder 
      BackColor       =   &H00F7FFF7&
      Height          =   330
      Left            =   116
      MaxLength       =   8
      TabIndex        =   2
      Top             =   2428
      Width           =   4725
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&X)"
      Height          =   495
      Left            =   4024
      Style           =   1  '�׷���
      TabIndex        =   6
      Top             =   3835
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00DBE6E6&
      Caption         =   "�� ��(&S)"
      Height          =   495
      Left            =   2809
      Style           =   1  '�׷���
      TabIndex        =   4
      Top             =   3835
      Width           =   1215
   End
   Begin VB.CommandButton cmdResult 
      BackColor       =   &H00DBE6E6&
      Caption         =   "..."
      Height          =   375
      Left            =   4856
      Style           =   1  '�׷���
      TabIndex        =   1
      Top             =   1470
      Width           =   390
   End
   Begin VB.CommandButton cmdOrder 
      BackColor       =   &H00DBE6E6&
      Caption         =   "..."
      Height          =   375
      Left            =   4856
      Style           =   1  '�׷���
      TabIndex        =   0
      Top             =   516
      Width           =   390
   End
   Begin MedControls1.LisLabel LisLabel2 
      Height          =   375
      Left            =   116
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   111
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
      Caption         =   "�� �������� �������"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel1 
      Height          =   375
      Left            =   116
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1047
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
      Caption         =   "�� ������� �������"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblOrder 
      Height          =   375
      Left            =   116
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   516
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
   End
   Begin MedControls1.LisLabel lblResult 
      Height          =   375
      Left            =   116
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1472
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel3 
      Height          =   375
      Left            =   109
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2920
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
      Caption         =   "�� ������ϸ� Ȯ����"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel4 
      Height          =   375
      Left            =   116
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1983
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
      Caption         =   "�� �������ϸ�"
      Appearance      =   0
   End
End
Attribute VB_Name = "frmIISConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIISConfig.frm
'   �ۼ���  : ������
'   ��  ��  : BEP2000 �ɼǼ�����
'   �ۼ���  : 2007-09-04
'   ��  ��  :
'-----------------------------------------------------------------------------'

Option Explicit

Private WithEvents mFolder1 As clsIISFolderSelect    '��������1
Attribute mFolder1.VB_VarHelpID = -1
Private WithEvents mFolder2 As clsIISFolderSelect    '��������2
Attribute mFolder2.VB_VarHelpID = -1

Private mEqpKey As String   '���Ű

Public Property Let EqpKey(ByVal vData As String)
    mEqpKey = vData
End Property

Private Sub Form_Load()
    If mEqpKey = "" Then
        MsgBox "���Ű�� �Էµ��� �ʾ����ϴ�.", vbInformation, "����"
        Unload Me
    End If
    
    lblOrder.Caption = mOrderPath
    lblResult.Caption = mResultPath
    txtOrder.Text = mOrderFileNm
    txtResult.Text = mResultFileNm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISConfig = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim strFileNm   As String   'INI���� ���+���ϸ�

    strFileNm = IniPath & "IIS.ini"
    
    mOrderPath = Trim$(lblOrder.Caption)
    mResultPath = Trim$(lblResult.Caption)
    mOrderFileNm = Trim$(txtOrder.Text)
    mResultFileNm = Trim$(txtResult.Text)
    
    Call mWriteINI(strFileNm, mEqpKey, "OrderPath", mOrderPath)
    Call mWriteINI(strFileNm, mEqpKey, "ResultPath", mResultPath)
    Call mWriteINI(strFileNm, mEqpKey, "OrderFileNm", mOrderFileNm)
    Call mWriteINI(strFileNm, mEqpKey, "ResultFileNm", mResultFileNm)
    
    MsgBox "���������� ����Ǿ����ϴ�.", vbInformation, "����"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOrder_Click()
    Set mFolder1 = New clsIISFolderSelect
    
    With mFolder1
        .Path = mOrderPath
        .Caption = "BEP2000 �������� �������"
        .ShowFolderSelect
    End With
    Set mFolder1 = Nothing
End Sub

Private Sub cmdResult_Click()
    Set mFolder2 = New clsIISFolderSelect
    
    With mFolder2
        .Path = mResultPath
        .Caption = "BEP2000 ������� �������"
        .ShowFolderSelect
    End With
    Set mFolder2 = Nothing
End Sub

Private Sub mFolder1_SelectedFolder(ByVal pSelFolder As String)
    lblOrder.Caption = pSelFolder
End Sub

Private Sub mFolder2_SelectedFolder(ByVal pSelFolder As String)
    lblResult.Caption = pSelFolder
End Sub

VERSION 5.00
Object = "{8120B504-1DBA-11D3-9D9C-00104B16DCF8}#3.0#0"; "MedControls1.ocx"
Begin VB.Form frmIISConfig 
   BackColor       =   &H00DBE6E6&
   BorderStyle     =   1  '���� ����
   Caption         =   "IT3000 ����"
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
   Begin VB.TextBox lblResult 
      Height          =   330
      Left            =   135
      TabIndex        =   19
      Top             =   1485
      Width           =   4695
   End
   Begin VB.TextBox lblOrder 
      Height          =   330
      Left            =   135
      TabIndex        =   18
      Top             =   540
      Width           =   4695
   End
   Begin VB.TextBox txtOrderSec 
      Alignment       =   2  '��� ����
      BackColor       =   &H00F7FFF7&
      Height          =   330
      Left            =   135
      MaxLength       =   8
      TabIndex        =   13
      Top             =   2475
      Width           =   2610
   End
   Begin VB.TextBox txtResultSec 
      Alignment       =   2  '��� ����
      BackColor       =   &H00F7FFF7&
      Height          =   330
      Left            =   135
      MaxLength       =   8
      TabIndex        =   12
      Top             =   3375
      Width           =   2610
   End
   Begin VB.TextBox txtResult 
      BackColor       =   &H00F7FFF7&
      Height          =   330
      Left            =   4395
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3340
      Visible         =   0   'False
      Width           =   4725
   End
   Begin VB.TextBox txtOrder 
      BackColor       =   &H00F7FFF7&
      Height          =   330
      Left            =   4395
      MaxLength       =   8
      TabIndex        =   2
      Top             =   2428
      Visible         =   0   'False
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
   Begin MedControls1.LisLabel lblOrder2 
      Height          =   375
      Left            =   116
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   516
      Visible         =   0   'False
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
   Begin MedControls1.LisLabel lblResult1 
      Height          =   375
      Left            =   116
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1472
      Visible         =   0   'False
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
      Left            =   4380
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2925
      Visible         =   0   'False
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
      Left            =   4395
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1980
      Visible         =   0   'False
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
   Begin MedControls1.LisLabel LisLabel5 
      Height          =   375
      Left            =   135
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2955
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
      Caption         =   "�� ������� ��ȸ�ֱ�"
      Appearance      =   0
   End
   Begin MedControls1.LisLabel LisLabel6 
      Height          =   375
      Left            =   135
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2025
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
      Caption         =   "�� �������� ��ȸ�ֱ�"
      Appearance      =   0
   End
   Begin VB.Label Label2 
      Caption         =   "sec"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2835
      TabIndex        =   17
      Top             =   3465
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "sec"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2835
      TabIndex        =   16
      Top             =   2565
      Width           =   600
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
'   ��  ��  : IT3000 �ɼǼ�����
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
    
    lblOrder.Text = mOrderPath
    lblResult.Text = mResultPath
    txtOrder.Text = mOrderFileNm
    txtResult.Text = mResultFileNm
    txtOrderSec.Text = mOrderRefresh
    txtResultSec.Text = mResultRefresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISConfig = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim strFileNm   As String   'INI���� ���+���ϸ�

    strFileNm = IniPath & "IIS.ini"
    
    mOrderPath = Trim$(lblOrder.Text)
    mResultPath = Trim$(lblResult.Text)
    mOrderFileNm = Trim$(txtOrder.Text)
    mResultFileNm = Trim$(txtResult.Text)
    mOrderRefresh = Trim$(txtOrderSec.Text)
    mResultRefresh = Trim$(txtResultSec.Text)
    
    Call mWriteINI(strFileNm, mEqpKey, "OrderPath", mOrderPath)
    Call mWriteINI(strFileNm, mEqpKey, "ResultPath", mResultPath)
    Call mWriteINI(strFileNm, mEqpKey, "OrderFileNm", mOrderFileNm)
    Call mWriteINI(strFileNm, mEqpKey, "ResultFileNm", mResultFileNm)
    Call mWriteINI(strFileNm, mEqpKey, "OrderRefresh", mOrderRefresh)
    Call mWriteINI(strFileNm, mEqpKey, "ResultRefresh", mResultRefresh)
    
    MsgBox "���������� ����Ǿ����ϴ�.", vbInformation, "����"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOrder_Click()
    Set mFolder1 = New clsIISFolderSelect
    
    With mFolder1
        .Path = mOrderPath
        .Caption = "IT3000 �������� �������"
        .ShowFolderSelect
    End With
    Set mFolder1 = Nothing
End Sub

Private Sub cmdResult_Click()
    Set mFolder2 = New clsIISFolderSelect
    
    With mFolder2
        .Path = mResultPath
        .Caption = "IT3000 ������� �������"
        .ShowFolderSelect
    End With
    Set mFolder2 = Nothing
End Sub

Private Sub mFolder1_SelectedFolder(ByVal pSelFolder As String)
    lblOrder.Text = pSelFolder
End Sub

Private Sub mFolder2_SelectedFolder(ByVal pSelFolder As String)
    lblResult.Text = pSelFolder
End Sub

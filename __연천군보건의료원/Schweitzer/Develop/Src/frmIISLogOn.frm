VERSION 5.00
Begin VB.Form frmIISLogOn 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  '���� ����
   ClientHeight    =   3930
   ClientLeft      =   1290
   ClientTop       =   780
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00A4BFC3&
   FillStyle       =   0  '�ܻ�
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3930
   ScaleWidth      =   6000
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F7F0F0&
      BorderStyle     =   0  '����
      ForeColor       =   &H00FFFFFF&
      Height          =   3990
      Left            =   0
      Picture         =   "frmIISLogOn.frx":0000
      ScaleHeight     =   3990
      ScaleWidth      =   6075
      TabIndex        =   4
      Top             =   0
      Width           =   6075
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00F7F3F8&
         Caption         =   "�� ��"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4755
         Style           =   1  '�׷���
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "128"
         Top             =   3210
         Width           =   900
      End
      Begin VB.CommandButton cmdConfirm 
         BackColor       =   &H00F7F3F8&
         Caption         =   "Ȯ ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3855
         Style           =   1  '�׷���
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "128"
         Top             =   3210
         Width           =   900
      End
      Begin VB.TextBox txtUserId 
         Alignment       =   2  '��� ����
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3885
         TabIndex        =   0
         Top             =   2010
         Width           =   1725
      End
      Begin VB.TextBox txtPass 
         Alignment       =   2  '��� ����
         BorderStyle     =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  '��� ����
         Left            =   3885
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   2790
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   300
         TabIndex        =   14
         Top             =   3315
         Width           =   2970
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00CDD19E&
         BorderWidth     =   3
         Height          =   330
         Left            =   3855
         Top             =   2760
         Width           =   1785
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00CDD19E&
         BorderWidth     =   3
         Height          =   330
         Left            =   3855
         Top             =   2370
         Width           =   1785
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00CDD19E&
         BorderWidth     =   3
         Height          =   330
         Left            =   3855
         Top             =   1980
         Width           =   1785
      End
      Begin VB.Label lblSysName 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Interface System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CDD19E&
         Height          =   555
         Index           =   0
         Left            =   1020
         TabIndex        =   12
         Top             =   915
         Width           =   3855
      End
      Begin VB.Label lblSysName 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Interface System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   2
         Left            =   1005
         TabIndex        =   13
         Top             =   900
         Width           =   3855
      End
      Begin VB.Label lblSysName 
         Alignment       =   2  '��� ����
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "Interface System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Index           =   1
         Left            =   1080
         TabIndex        =   11
         Top             =   945
         Width           =   3855
      End
      Begin VB.Label lblName 
         Alignment       =   2  '��� ����
         Appearance      =   0  '���
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '����
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3855
         TabIndex        =   8
         Top             =   2430
         Width           =   1785
      End
      Begin VB.Label lblUser 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "�����ID"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2715
         TabIndex        =   7
         Top             =   2070
         Width           =   1020
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "��й�ȣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2760
         TabIndex        =   6
         Top             =   2835
         Width           =   990
      End
      Begin VB.Label lblUserNm 
         Alignment       =   1  '������ ����
         BackStyle       =   0  '����
         Caption         =   "����ڸ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   2715
         TabIndex        =   5
         Top             =   2445
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F1F2E3&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3855
         TabIndex        =   9
         Top             =   2385
         Visible         =   0   'False
         Width           =   1785
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2430
      TabIndex        =   10
      Top             =   1755
      Width           =   1215
   End
End
Attribute VB_Name = "frmIISLogOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   ���ϸ�  : frmIISLogOn.frm
'   �ۼ���  :
'   ��  ��  : �α�����
'   ��  ��  :
'-----------------------------------------------------------------------------'
Option Explicit

Private mLogOn      As clsIISLogOn      '�α��� Ŭ����
Private mIsLogOn    As Boolean          'True(�α��� ����), Flase(����)

Public Property Get IsLogOn() As Boolean
    IsLogOn = mIsLogOn
End Property

Private Sub Form_Load()
    Set mLogOn = New clsIISLogOn
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mLogOn = Nothing
    Set frmIISLogOn = Nothing
End Sub

Private Sub cmdConfirm_Click()
    If txtPass.Text = "" Then
        MsgBox "��й�ȣ�� �Է��ϼ���.", vbInformation, "����"
        Call txtPass_GotFocus
        Exit Sub
    End If

    If Trim(txtPass.Text) = mLogOn.LoginPass Then
        Call SetUserInfo(mLogOn.EMPID, mLogOn.EMPNM)
        mIsLogOn = True
        Unload Me
    Else
        MsgBox "��й�ȣ�� Ʋ���ϴ�. ��й�ȣ�� Ȯ���ϼ���.", vbInformation, "����"
        Call txtPass_GotFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    mIsLogOn = False
    Unload Me
End Sub

Private Sub txtUserId_Change()
    lblName.Caption = ""
End Sub

Private Sub txtUserId_GotFocus()
    With txtUserId
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    cmdConfirm.Enabled = False
End Sub

Private Sub txtUserId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtUserId_Validate(Cancel As Boolean)
    If CheckId Then
        Cancel = False
        cmdConfirm.Enabled = True
    Else
        Cancel = True
        Call txtUserId_GotFocus
        cmdConfirm.Enabled = False
    End If
End Sub

Private Sub txtPass_GotFocus()
    With txtPass
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdConfirm_Click
    End If
End Sub

'-----------------------------------------------------------------------------'
'   ��� : ���̵��� ��ȿ�� �˻�
'   ��ȯ : True(��ȿ), Flase(��ȿ)
'-----------------------------------------------------------------------------'
Private Function CheckId() As Boolean
    If txtUserId.Text = "" Then
        MsgBox "�α��� ���̵� �Է��ϼ���.", vbInformation, "����"
        CheckId = False
        Exit Function
    End If

    If mLogOn.GetEmpInfo(Trim(txtUserId.Text)) = False Then
        MsgBox "��ϵ��� ���� ID�Դϴ�. �α��� ID�� Ȯ���ϼ���.", vbInformation, "����"
        CheckId = False
    Else
        lblName.Caption = mLogOn.EMPNM
        CheckId = True
    End If
End Function

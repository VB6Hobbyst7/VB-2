VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmPassChange 
   Caption         =   "�۾��� ��й�ȣ ����"
   ClientHeight    =   2805
   ClientLeft      =   4965
   ClientTop       =   2625
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   4560
   Begin Threed.SSPanel SSPanel3 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2340
      Width           =   4320
      _Version        =   65536
      _ExtentX        =   7620
      _ExtentY        =   661
      _StockProps     =   15
      ForeColor       =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1650
      Left            =   180
      TabIndex        =   7
      Top             =   60
      Width           =   4185
      _Version        =   65536
      _ExtentX        =   7382
      _ExtentY        =   2910
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TxtPass3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         IMEMode         =   3  '��� ����
         Left            =   1770
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1200
         Width           =   2100
      End
      Begin VB.TextBox TxtPass2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  '��� ����
         Left            =   1770
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   2100
      End
      Begin VB.TextBox TxtPass1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         IMEMode         =   3  '��� ����
         Left            =   1785
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   480
         Width           =   2100
      End
      Begin VB.TextBox TxtSabun 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1785
         TabIndex        =   0
         Top             =   120
         Width           =   2100
      End
      Begin VB.Label LabelPassward 
         Appearance      =   0  '���
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  '����
         Caption         =   "��й�ȣȮ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1305
         Width           =   1350
      End
      Begin VB.Label LabelPassward 
         Appearance      =   0  '���
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  '����
         Caption         =   "�����й�ȣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label LabelIdnumber 
         Appearance      =   0  '���
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  '����
         Caption         =   "��          ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   255
         TabIndex        =   9
         Top             =   195
         Width           =   1305
      End
      Begin VB.Label LabelPassward 
         Appearance      =   0  '���
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  '����
         Caption         =   "�� ��   �� ȣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   8
         Top             =   570
         Width           =   1290
      End
   End
   Begin Threed.SSCommand CmdCancel 
      Height          =   420
      Left            =   2100
      TabIndex        =   5
      Top             =   1800
      Width           =   2310
      _Version        =   65536
      _ExtentX        =   4075
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "��й�ȣ �������[&C]"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdOk 
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1950
      _Version        =   65536
      _ExtentX        =   3440
      _ExtentY        =   741
      _StockProps     =   78
      Caption         =   "��й�ȣ ����(&O)"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmPassChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FnPassCount     As Integer '��й�ȣ ����Ƚ��
Dim FnSabunCnt      As Integer '��� ���� Ƚ��

Dim FstrPassWord    As String
Dim FstrPassName    As String
Dim FstrPassGrade   As String
Dim FstrPassPart    As String * 1
Dim FstrPassCharge  As String


Private Sub CmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub CmdOK_Click()
    
    If Trim(TxtPass2.Text) = "" Then
        MsgBox "�����Ͻ� ��й�ȣ�� �����Դϴ�.", vbCritical, "��й�ȣ ����"
        TxtPass2.SetFocus
        Exit Sub
    End If
    
    If TxtPass2.Text <> TxtPass3.Text Then
        MsgBox "�����Ͻ� ��й�ȣ�� ��Ȯ���� ����", vbCritical, "����"
        TxtPass3.SetFocus
        Exit Sub
    End If
    
    '����� ��й�ȣ�� ����
    strSQL = "         UPDATE TW_MIS_PMPA.TWBAS_PASS "
    strSQL = strSQL & "   SET PassWord = '" & TxtPass2.Text & "'                "
    strSQL = strSQL & " WHERE IDNumber = " & Val(TxtSabun.Text)
    Result = AdoExecute(strSQL)
    
    Unload Me

End Sub

Private Sub Form_Load()
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    TxtSabun.Text = ""
    TxtPass1.Text = "":    TxtPass2.Text = ""
    TxtPass3.Text = ""
    FnPassCount = 0:       FnSabunCnt = 0
    If GnJobSabun <> 0 Then
        TxtSabun.Text = GnJobSabun
        Call READ_BAS_Pass
        TxtSabun.Enabled = False
    End If
    TxtPass2.Enabled = False
    TxtPass3.Enabled = False
    CmdOK.Enabled = False
    
'H    Call FORM_CENTER(Me)
    
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub TxtPass1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub TxtPass1_LostFocus()
    Dim strPassWord     As String
    
    If Trim(TxtPass1.Text) = "" Then Exit Sub
    
    FnPassCount = FnPassCount + 1
    If FnPassCount > 2 Then
        GstrMsgPromPt = "ID�� ��й�ȣ�� Ȯ���� �ٽ� �۾��� �Ͻʽÿ�." & Chr(13) & Chr(13)
        GstrMsgPromPt = GstrMsgPromPt & "�ܷ� OCS ���α׷��� ���� �մϴ�."
        MsgBox GstrMsgPromPt, vbCritical + vbExclamation, "�۾�����"
        End
    End If
    
    '����� Check
    If TxtSabun.Text = "" Then
        MsgBox "����� �����Դϴ�.", , "Ȯ��"
        TxtSabun.SetFocus
        Exit Sub
    End If
    
    '�Է��� �̹й�ȣ�� Check
    strPassWord = UCase(Trim(TxtPass1.Text))
    If strPassWord <> FstrPassWord Then
        MsgBox "��Ȯ�� ��й�ȣ�� �Է��ϼ���", , "Ȯ��"
        TxtPass1.Text = ""
        TxtPass1.SetFocus
        Exit Sub
    End If
    
    TxtSabun.Enabled = False
    TxtPass1.Enabled = False
    TxtPass2.Enabled = True
    TxtPass3.Enabled = True
    CmdOK.Enabled = True
    
    TxtPass2.SetFocus

End Sub

Private Sub TxtPass2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub TxtPass2_LostFocus()
    Dim i           As Integer
    Dim strChar     As String * 1
    
    TxtPass2.Text = Trim(UCase(TxtPass2.Text))
    If TxtPass2.Text = "" Then Exit Sub
    
    If Len(TxtPass2.Text) < 2 Then
        MsgBox "��й�ȣ�� �ݵ�� 2�ڸ� �̻� ������", vbCritical, "����"
        TxtPass2.SetFocus
        Exit Sub
    End If
    
    '�������� Check
    For i = 1 To Len(TxtPass2.Text)
        strChar = Mid(TxtPass2.Text, i, 1)
        Select Case strChar
            Case "0" To "9":
            Case "A" To "Z":
            Case Else:
                 MsgBox "��й�ȣ�� 0-9,A-Z ���ڸ� ����� ������", , "Ȯ��"
                 TxtPass2.Text = ""
                 TxtPass2.SetFocus
                 Exit Sub
        End Select
    Next i
    
End Sub

Private Sub TxtPass3_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{Tab}"

End Sub

Private Sub TxtPass3_LostFocus()
    
    TxtPass3.Text = Trim(UCase(TxtPass3.Text))

End Sub

Private Sub TxtSabun_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub TxtSabun_LostFocus()
    
    TxtSabun.Text = Trim(TxtSabun.Text)
    
    If TxtSabun.Text = "" Then Exit Sub
    FstrPassName = ""
    
    '����� ������ 3���̻��̸� ������ ������
    FnSabunCnt = FnSabunCnt + 1
    If FnSabunCnt > 3 Then
        GstrMsgPromPt = "��Ȯ�� ����� Ȯ����" & vbCrLf
        GstrMsgPromPt = GstrMsgPromot & "�ٽ� �۾��� �Ͻʽÿ�."
        MsgBox GstrMsgPromPt, vbCritical + vbExclamation, "�۾�����"
        
        Call DbAdoDisConnect
        End
    End If
    
    Call READ_BAS_Pass
    
    If FstrPassName = "" Then  '�̸��� �����̸�
        GstrMsgPromPt = "����� ��ϵ��� �ʾҽ��ϴ�." & vbCrLf
        GstrMsgPromPt = GstrMsgPromPt & "����� Ȯ���� �����ø�" & vbCrLf
        GstrMsgPromPt = GstrMsgPromPt & "����ǿ� ���� �ֽʽÿ�."
        MsgBox GstrMsgPromPt, vbCritical + vbExclamation, "Ȯ��"
        TxtSabun.SetFocus
        Exit Sub
    End If
    
End Sub
Sub READ_BAS_Pass()
       
    
    'BAS_PASS���� �ش� ����� ��ϵǾ����� ����
    strSQL = "         SELECT Name, PassWord, Grade, Part, ProgramID             " 'change
    strSQL = strSQL & "  FROM  TW_MIS_PMPA.TWBAS_PASS "
    strSQL = strSQL & " WHERE  ProgramID = ' '                                   "
    strSQL = strSQL & "   AND  IDNumber  = " & Val(TxtSabun.Text)
    Result = AdoOpenSet(ADORES, strSQL)
    
    If Not ADORES.EOF Then
        FstrPassWord = Trim(AdoGetString(ADORES, "PassWord"))
        FstrPassName = Trim(AdoGetString(ADORES, "Name"))
        FstrPassGrade = Trim(AdoGetString(ADORES, "Grade"))
        FstrPassPart = AdoGetString(ADORES, "Part")
'H        FstrPassCharge = Trim(AdoGetString(RS1, "Charge"))
    Else
        FstrPassWord = "":      FstrPassName = ""
        FstrPassGrade = "":     FstrPassPart = ""
'H        FstrPassCharge = ""
    End If
        
'H    Call AdoCloseSet(AdoRes)

End Sub

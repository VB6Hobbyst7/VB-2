VERSION 5.00
Begin VB.Form frmDbConfig 
   Caption         =   "ȯ�� ����"
   ClientHeight    =   5850
   ClientLeft      =   3300
   ClientTop       =   2700
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   6525
   Begin VB.TextBox txCauUsn 
      Alignment       =   1  '������ ����
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   3600
      TabIndex        =   27
      Text            =   "40"
      Top             =   4470
      Width           =   975
   End
   Begin VB.Frame DTScFrame 
      Caption         =   "DB Info"
      Height          =   1335
      Left            =   240
      TabIndex        =   19
      Top             =   600
      Width           =   6135
      Begin VB.CommandButton cmdScDBConTest 
         Caption         =   "�����׽�Ʈ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   0
         Left            =   4425
         TabIndex        =   23
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox txtScDTDataSc 
         BackColor       =   &H80000006&
         ForeColor       =   &H80000005&
         Height          =   270
         Left            =   1410
         TabIndex        =   22
         Text            =   "KORAOBS01"
         Top             =   900
         Width           =   2895
      End
      Begin VB.TextBox txtScDTPW 
         BackColor       =   &H80000006&
         ForeColor       =   &H80000005&
         Height          =   270
         IMEMode         =   3  '��� ����
         Left            =   1410
         PasswordChar    =   "*"
         TabIndex        =   21
         Text            =   "monitering"
         Top             =   570
         Width           =   2895
      End
      Begin VB.TextBox txtScDTID 
         BackColor       =   &H80000006&
         ForeColor       =   &H80000005&
         Height          =   270
         Left            =   1410
         TabIndex        =   20
         Text            =   "monitering"
         Top             =   225
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "�����ͼҽ�:"
         Height          =   180
         Left            =   375
         TabIndex        =   26
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "��й�ȣ:"
         Height          =   180
         Left            =   555
         TabIndex        =   25
         Top             =   600
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "�����:"
         Height          =   180
         Left            =   735
         TabIndex        =   24
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.TextBox txCauRt 
      Alignment       =   1  '������ ����
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   3600
      TabIndex        =   13
      Text            =   "40"
      Top             =   4155
      Width           =   975
   End
   Begin VB.TextBox txCauAg 
      Alignment       =   1  '������ ����
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   3600
      TabIndex        =   12
      Text            =   "60"
      Top             =   3795
      Width           =   975
   End
   Begin VB.TextBox txCauTw 
      Alignment       =   1  '������ ����
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   3585
      TabIndex        =   11
      Text            =   "60"
      Top             =   3435
      Width           =   975
   End
   Begin VB.TextBox txCauDtCdma 
      Alignment       =   1  '������ ����
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   270
      Left            =   3585
      TabIndex        =   10
      Text            =   "30"
      Top             =   3075
      Width           =   975
   End
   Begin VB.TextBox txCauDtVpn 
      Alignment       =   1  '������ ����
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   270
      IMEMode         =   7  '���� ����
      Left            =   3585
      TabIndex        =   9
      Text            =   "5"
      Top             =   2715
      Width           =   975
   End
   Begin VB.CommandButton btnCgSave 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4635
      TabIndex        =   1
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton btnCgCancel 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5595
      TabIndex        =   0
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  '������ ����
      Caption         =   "U S N : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1530
      TabIndex        =   29
      Top             =   4470
      Width           =   2070
   End
   Begin VB.Label Label7 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   4665
      TabIndex        =   28
      Top             =   4470
      Width           =   180
   End
   Begin VB.Label Label24 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   4665
      TabIndex        =   18
      Top             =   4155
      Width           =   180
   End
   Begin VB.Label Label23 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   4665
      TabIndex        =   17
      Top             =   3795
      Width           =   180
   End
   Begin VB.Label Label22 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   4665
      TabIndex        =   16
      Top             =   3435
      Width           =   180
   End
   Begin VB.Label Label17 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   4665
      TabIndex        =   15
      Top             =   3075
      Width           =   180
   End
   Begin VB.Label Label18 
      Alignment       =   2  '��� ����
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   4665
      TabIndex        =   14
      Top             =   2715
      Width           =   180
   End
   Begin VB.Label Label16 
      Alignment       =   1  '������ ����
      Caption         =   "�ؾ������ : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1530
      TabIndex        =   8
      Top             =   4170
      Width           =   2070
   End
   Begin VB.Label Label12 
      Alignment       =   1  '������ ����
      Caption         =   "�ؼ��������� : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1530
      TabIndex        =   7
      Top             =   3825
      Width           =   2070
   End
   Begin VB.Label Label11 
      Alignment       =   1  '������ ����
      Caption         =   "�����ؾ�������� : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1545
      TabIndex        =   6
      Top             =   3450
      Width           =   2070
   End
   Begin VB.Label Label10 
      Alignment       =   1  '������ ����
      Caption         =   "����������(CDMA) : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1545
      TabIndex        =   5
      Top             =   3105
      Width           =   2070
   End
   Begin VB.Label Label6 
      Alignment       =   1  '������ ����
      Caption         =   "����������(VPN) : "
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1545
      TabIndex        =   4
      Top             =   2715
      Width           =   2070
   End
   Begin VB.Label Label5 
      Alignment       =   2  '��� ����
      Caption         =   "[ ��� �ð� ���� ]"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   2190
      Width           =   6375
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��� ����
      Caption         =   "[ DataBase���� ���� ���� ]"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmDbConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCgCancel_Click()
    Unload Me
End Sub

Private Sub btnCgSave_Click()
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'" 1. ����       : ���밳����
'" 2. ���������� : ȯ�漳�� ���� ����
'" 5. �ۼ���     : ��â��
'" 6. �ۼ���     : 2008/09/17
'" 7. ���ϰ�     :
'" 8. ���� �̷�  :
'"
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
On Error GoTo ErrorHandler

    Dim Fnum As Long

    If Trim(txtScDTID.Text) = "" Then
        MsgBox "DB ���� ID�� �Է����ּ���."
        Exit Sub
    ElseIf Trim(txtScDTPW.Text) = "" Then
        MsgBox "DB ���� �н����带 �Է����ּ���."
        Exit Sub
    ElseIf Trim(txtScDTDataSc.Text) = "" Then
        MsgBox "DB ���� DataSource������ �Է����ּ���."
        Exit Sub
    End If

        
    If Not IsNumeric(txCauDtVpn.Text) Then
        MsgBox "���ڸ� �Է����ּ���."
        txCauDtVpn.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txCauDtCdma.Text) Then
        MsgBox "���ڸ� �Է����ּ���."
        txCauDtCdma.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txCauTw.Text) Then
        MsgBox "���ڸ� �Է����ּ���."
        txCauTw.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txCauAg.Text) Then
        MsgBox "���ڸ� �Է����ּ���."
        txCauAg.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txCauRt.Text) Then
        MsgBox "���ڸ� �Է����ּ���."
        txCauRt.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txCauUsn.Text) Then
        MsgBox "���ڸ� �Է����ּ���."
        txCauUsn.SetFocus
        Exit Sub
        
    End If
    
    Fnum = FreeFile

    Open App.Path & "\Config.bin" For Output As #Fnum
        Print #Fnum, "[DataBaseInfo]"
        Print #Fnum, txtScDTID & "," & txtScDTPW & "," & txtScDTDataSc

        
        Print #Fnum, "[CAUTION]"
        Print #Fnum, "DT_VPN=" & txCauDtVpn
        Print #Fnum, "DT_CDMA=" & txCauDtCdma
        Print #Fnum, "TW=" & txCauTw
        Print #Fnum, "AG=" & txCauAg
        Print #Fnum, "RT=" & txCauRt
        Print #Fnum, "USN=" & txCauUsn

    Close #Fnum
    
    

    Call MsgBox("ȯ�漳�� ������ �����Ͽ����ϴ�.", vbInformation + vbOKOnly, mainFrm.Caption)
    
    'ȯ�漳�� �� Load
    mainFrm.GetConfig
    
    Unload frmDbConfig
    
    Exit Sub
ErrorHandler:
    If Err.Number <> 0 Then
        Call LogWrite("btnCgSave_Click : " & Err.Number & "-" & Err.Description)
        Err.Clear
        Call MsgBox("���������� �����Ҽ� �����ϴ�.", vbCritical + vbOKOnly, mainFrm.Caption)
        Exit Sub
    End If
End Sub

Private Sub cmdScDBConTest_Click(Index As Integer)

On Error GoTo ErrorHandler
    
    Dim Conn As ADODB.Connection
    Dim bOpen As Boolean
    Dim strId, strPw, strSource As String
    
    If Index = 0 Then
        strId = txtScDTID
        strPw = txtScDTPW
        strSource = txtScDTDataSc
    Else
        MsgBox "���ǵ��� ���� �ε����� �ԷµǾ����ϴ�."
    End If
    
   
    bOpen = False
    
         ConnectionString = "Provider=MSDAORA.1;Password=" + strPw + ";User ID=" + strId + ";Data Source=" + strSource + ";Persist Security Info=True"
    
    
    Set Conn = New ADODB.Connection
    Conn.Open ConnectionString
    
    If Conn.State = adStateOpen Then
        Call MsgBox("�����ͺ��̽� ���� �׽��׿� �����Ͽ����ϴ�.", vbInformation + vbOKOnly, mainFrm.Caption)
        bOpen = True
    Else
        Call MsgBox("�����ͺ��̽� ���� �׽��׿� �����Ͽ����ϴ�.", vbCritical + vbOKOnly, mainFrm.Caption)
        bOpen = False
    End If
    
    
    If bOpen = True Then
        If Conn.State = adStateOpen Then
            Conn.Close
        End If
    End If

    If Not Conn Is Nothing Then
        Set Conn = Nothing
    End If

    Exit Sub
ErrorHandler:
    If Err.Number <> 0 Then
        Err.Clear
        Call MsgBox("�����ͺ��̽� ���� �׽��׿� �����Ͽ����ϴ�.", vbCritical + vbOKOnly, mainFrm.Caption)
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    'DB Connection Info
    txtScDTID.Text = CfgDb.ID
    txtScDTPW.Text = CfgDb.PW
    txtScDTDataSc.Text = CfgDb.DataSource
        
    txCauDtVpn.Text = strJowiVPNCautionMin
    txCauDtCdma.Text = strJowiCDMACautionMin
    txCauTw.Text = strTwCautionMin
    txCauAg.Text = strAgCautionMin
    txCauRt.Text = strRtCautionMin
    txCauUsn.Text = strUsnCautionMin
End Sub

VERSION 5.00
Begin VB.Form frmUser 
   Caption         =   "�����"
   ClientHeight    =   1860
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox txtPW 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  '��� ����
      Left            =   1590
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   2025
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "Ȯ��"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1110
      Width           =   975
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1590
      TabIndex        =   0
      Top             =   150
      Width           =   2025
   End
   Begin VB.Label lblErr 
      BackStyle       =   0  '����
      Caption         =   "* ����� ID�� Password �� �߸��Ǿ����ϴ�."
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   90
      TabIndex        =   5
      Top             =   1590
      Width           =   4635
   End
   Begin VB.Label Label2 
      Alignment       =   1  '������ ����
      Caption         =   "�н����� :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   3
      Top             =   630
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  '������ ����
      Caption         =   "���̵� :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   2
      Top             =   180
      Width           =   1125
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUser_Click()
    If Trim(txtUser.Text) = "" Then
        lblErr = "* ����� ���̵� �Է��ϼ���."
        txtUser.SetFocus
        Exit Sub
    End If
    
    If Trim(txtPW.Text) = "" Then
        lblErr = "* ��й�ȣ�� �Է��ϼ���."
        txtPW.SetFocus
        Exit Sub
    End If
    
'    lsWK = Get_WKID(Trim(txtID.Text))
    gIFName = ""
    Online_TLA gXml_S24, Trim(txtUser.Text), Trim(txtPW.Text)
    If gIFName = "" Then
        lblErr = "* ���̵�� �н����尡 ��ġ���� �ʽ��ϴ�."
        txtPW.Text = ""
        txtUser.Text = ""
        txtUser.SetFocus
    Else
        lblErr = ""
        frmInterface.lblUser.Caption = gIFName
        gIFUser = Trim(txtUser.Text)
'        frmInterface.Show 0
        Unload Me
    End If
 
End Sub

Private Sub Form_Load()
    lblErr.Caption = ""
    
End Sub


Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtPW.Text) = "" Then
            lblErr = "* ��й�ȣ�� �Է��ϼ���."
            txtPW.SetFocus
            Exit Sub
        Else
            lblErr = ""
            cmdUser_Click
            
        End If
        
    End If
    
    
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lsWK As Integer

    If KeyCode = 13 Then
        If Trim(txtUser.Text) = "" Then
            lblErr = "* ����� ���̵� �Է��ϼ���."
            txtUser.SetFocus
            Exit Sub
        Else
            txtPW.SetFocus
        End If

    End If
End Sub

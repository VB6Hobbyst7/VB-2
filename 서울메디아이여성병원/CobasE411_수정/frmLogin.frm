VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "Login"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7695
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox txttemp 
      Height          =   270
      Left            =   750
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
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
      Height          =   330
      IMEMode         =   3  '��� ����
      Left            =   5460
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2130
      Width           =   2025
   End
   Begin VB.TextBox txtID 
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
      Height          =   330
      Left            =   5460
      TabIndex        =   3
      Top             =   1680
      Width           =   2025
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   0
      Picture         =   "frmLogin.frx":058A
      Top             =   30
      Width           =   4470
   End
   Begin VB.Label lblErr 
      BackStyle       =   0  '����
      Caption         =   "* ����� ID�� Password �� �߸��Ǿ����ϴ�."
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   180
      TabIndex        =   7
      Top             =   2820
      Width           =   4635
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   6840
      TabIndex        =   6
      Top             =   2790
      Width           =   645
   End
   Begin VB.Label lblCommit 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Ȯ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   6000
      TabIndex        =   5
      Top             =   2790
      Width           =   645
   End
   Begin VB.Label lblPW 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "��й�ȣ :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4200
      TabIndex        =   2
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblID 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "���̵� :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   1740
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00FF8080&
      FillColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   4470
      Top             =   0
      Width           =   45
   End
   Begin VB.Label lblEquipName 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "CobasE411 Interface"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   4620
      TabIndex        =   0
      Top             =   930
      Width           =   6015
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    lblErr = ""
    GetLSetup
End Sub

Private Sub lblCancel_Click()
    Unload Me
End Sub

Private Sub lblCommit_Click()
    Dim lsWK As Boolean
    
    If Trim(txtID.Text) = "" Then
        lblErr = "* ����� ���̵� �Է��ϼ���."
        txtID.SetFocus
        Exit Sub
    End If

    If Trim(txtPW.Text) = "" Then
        lblErr = "* ��й�ȣ�� �Է��ϼ���."
        txtPW.SetFocus
        Exit Sub
    End If
    
    If Not Connect_Server Then
        MsgBox "������� �ʾҽ��ϴ�."
        Exit Sub
    End If
    
    SQL = "select usnm from mstusid where usid = '" & Trim(txtID.Text) & "' and pass = '" & Trim(txtPW.Text) & "'"
    
    res = db_select_Col(gServer, SQL)
    
    DisConnect_Server
    
    If res = 1 Then
        lblErr = ""
        gUserID = Trim(txtID.Text)
        
        Me.Hide
        frmInterface.txtUID.Text = Trim(gReadBuf(0))
        frmInterface.Show 0
        
    Else
        lblErr = "* �߸��� ����� �����Դϴ�."
        txtID.Text = ""
        txtPW.Text = ""
        txtID.SetFocus
    End If
    

End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lsWK As Integer

    If KeyCode = 13 Then
        If Trim(txtID.Text) = "" Then
            lblErr = "* ����� ���̵� �Է��ϼ���."
            txtID.SetFocus
            Exit Sub
        Else

                lblErr = ""
                txtPW.SetFocus

        End If
    End If
    
End Sub

Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtPW.Text) = "" Then
            lblErr = "* ��й�ȣ�� �Է��ϼ���."
            txtPW.SetFocus
            Exit Sub
        Else
            lblErr = ""
            lblCommit_Click
            
        End If
        
    End If
End Sub

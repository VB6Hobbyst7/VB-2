VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "Login"
   ClientHeight    =   3375
   ClientLeft      =   3240
   ClientTop       =   2925
   ClientWidth     =   6255
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6255
   Begin VB.TextBox txtTemp 
      Height          =   495
      Left            =   -1170
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   1365
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
      Left            =   4140
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2190
      Width           =   1575
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
      Left            =   4140
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   465
      Left            =   750
      Top             =   1020
      Width           =   105
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "���ִ��к���"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   525
      Left            =   1050
      TabIndex        =   10
      Top             =   210
      Width           =   2325
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H008080FF&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H008080FF&
      FillColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   90
      Top             =   2160
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H0080FFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   30
      Top             =   2130
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "���ܰ˻����а�"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3990
      TabIndex        =   8
      Top             =   360
      Width           =   3915
   End
   Begin VB.Label lblErr 
      BackStyle       =   0  '����
      Caption         =   "* ����� ID�� Password �� �߸��Ǿ����ϴ�."
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   210
      TabIndex        =   7
      Top             =   2790
      Width           =   3765
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "���"
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
      Left            =   5130
      TabIndex        =   6
      Top             =   2700
      Width           =   645
   End
   Begin VB.Label lblCommit 
      Alignment       =   2  '��� ����
      BackStyle       =   0  '����
      Caption         =   "Ȯ��"
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
      Left            =   4350
      TabIndex        =   5
      Top             =   2700
      Width           =   645
   End
   Begin VB.Label lblPW 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "��й�ȣ :"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblID 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "���̵� :"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   1830
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00FF8080&
      FillColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   -30
      Top             =   2130
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblEquipName 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "Cobas Taqman Interface"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   1020
      Width           =   4905
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
    GetSetup
'    Init_WK
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Unload frmInterface
    
End Sub

Private Sub lblCancel_Click()

    Unload Me
    
End Sub

Private Sub lblCommit_Click()
Dim lsWK As Integer

    If Trim(txtID.Text) = "" Then
        lblErr = "* ����� ���̵� �Է��ϼ���."
        txtID.SetFocus
        Exit Sub
    End If
    
'''    If Trim(txtPW.Text) = "" Then
'''        lblErr = "* ��й�ȣ�� �Է��ϼ���."
'''        txtPW.SetFocus
'''        Exit Sub
'''    End If
    
'    lsWK = Get_WKID(Trim(txtID.Text))
    'gIFName = ""
    If Not Connect_Server Then
        MsgBox "������� �ʾҽ��ϴ�."
        Exit Sub
    End If
    'gIFName = ""
    
'''    SQL = "SELECT OMT13NAME FROM VWOMTEMP13 WHERE OMT13EMPNO = '" & Trim(txtID.Text) & "' AND OMT13PASSWD = '" & Trim(txtPW.Text) & "'"
    
    SQL = "select netc1.fn_cs_login_password_check('" & txtID & "','" & txtPW & "') FROM dual"
    res = db_select_Col(gServer, SQL)
    
    cn_Ser.Close
    
    If Trim(gReadBuf(0)) = "False" Then
        lblErr = "*  ��й�ȣ�� Ȯ���� �ּ���."
        txtPW.Text = ""
        'txtID.Text = ""
        txtPW.SetFocus
    ElseIf Trim(gReadBuf(0)) = "True" Then
        lblErr = ""
        gExamUID = Trim(txtID.Text)
        
        frmInterface.Show 0
        frmInterface.txtUID.Text = frmInterface
        Unload Me
    Else
        lblErr = "* ���̵� �� Ȯ���� �ּ���."
        txtPW.Text = ""
        'txtID.Text = ""
        txtID.SetFocus
    End If
    
    
    
'    If Trim(gWorker_Info.WK_PW) = Trim(txtPW.Text) And Trim(gWorker_Info.WK_ID) = Trim(txtID.Text) Then
'        lblErr = ""
'        frmInterface.lblUser.Caption = "����� : " & gWorker_Info.WK_NM
'        frmInterface.Show 0
'        Me.Hide
'
'    Else
'        lblErr = "* ��й�ȣ�� Ȯ���ϼ���."
'        txtPW.Text = ""
'        txtPW.SetFocus
'    End If
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
            'lblCommit_Click
            
'''            lblCommit
        End If
'            lsWK = Get_WKID(Trim(txtID.Text))
'            If lsWK > 0 Then
'                lblErr = ""
'                txtPW.SetFocus
'
'            Else
'                lblErr = "* �������� �ʴ� ���̵��Դϴ�."
'                txtID.Text = ""
'                txtID.SetFocus
'                Exit Sub
'            End If
'        End If
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

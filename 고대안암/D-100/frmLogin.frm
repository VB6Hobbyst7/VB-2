VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "Login"
   ClientHeight    =   3780
   ClientLeft      =   3240
   ClientTop       =   2925
   ClientWidth     =   6645
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6645
   Begin VB.TextBox txtLocate 
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
      Left            =   4470
      TabIndex        =   11
      Text            =   "AA"
      Top             =   2700
      Width           =   1575
   End
   Begin VB.TextBox txtTemp 
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   2400
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
      Left            =   4470
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2250
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
      Left            =   4470
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "�����ڵ� :"
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
      Left            =   3210
      TabIndex        =   12
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   465
      Left            =   1830
      Top             =   1050
      Width           =   105
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "���ȾϺ���"
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
      Left            =   510
      TabIndex        =   10
      Top             =   210
      Width           =   2745
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
      Left            =   2880
      TabIndex        =   8
      Top             =   360
      Width           =   3465
   End
   Begin VB.Label lblErr 
      BackStyle       =   0  '����
      Caption         =   "* ����� ID�� Password �� �߸��Ǿ����ϴ�."
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   210
      TabIndex        =   7
      Top             =   3330
      Width           =   4515
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
      Left            =   5460
      TabIndex        =   6
      Top             =   3240
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
      Left            =   4680
      TabIndex        =   5
      Top             =   3240
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
      Left            =   3210
      TabIndex        =   2
      Top             =   2220
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
      Left            =   3210
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
      Caption         =   "D100 Interface"
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
      Left            =   2070
      TabIndex        =   0
      Top             =   1050
      Width           =   4065
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
'''
'''    txtID = gTransUser
'''    txtPW = gTransUserPW
'''    txtID = "1952171"
'''    txtPW = "suwanmania"
    

    
End Sub

Private Sub lblCancel_Click()

    Unload Me
    
End Sub

Private Sub lblCommit_Click()
Dim lsWK As Integer
Dim vReturnValue As Variant
Dim iResCnt As Long
Dim strUserName As String
Dim iRes As Integer
Dim iiRes As Integer
Dim vRes As Variant


'''    txtID.Text = "13039"
'''    txtPW.Text = "rheo131313"
    
    gIFName = ""
    
    
    If Len(txtID.Text) < 1 Then
        lblErr = "* ����� ���̵� �Է��ϼ���."
        txtID.SetFocus
        Exit Sub
    End If
    
    If Len(txtPW.Text) < 1 Then
        lblErr = "* ��й�ȣ�� �Է��ϼ���."
        txtPW.SetFocus
        Exit Sub
    End If
    
    
    sEMRUser = "kuh_test"
    sEMRID = "tuxedo"
    sEMRPW = "01"
    
    iRes = TuxedoInit(sEMRUser, sEMRID, sEMRPW)
'''    MsgBox "1"
    
    iiRes = UserChk(txtID.Text, txtPW.Text, txtLocate.Text, vRes)
'''    MsgBox "2"
    
    Save_Raw_Data CStr(iiRes)
    iRes = TuxedoTerm
    
    
'''    Exit Sub
    
    If iiRes > 0 Then
        strUserName = vRes(0)
    Else
        strUserName = ""
    End If

'''    MsgBox strUserName
    
    gIFName = strUserName
'''    Online_TLA gXml_S24, Trim(txtID.text), Trim(txtPW.text)
    
'''    gTMAX.TP_INIT
'''    iResCnt = UserChk(txtID.Text, txtPW.Text, vReturnValue)
'''    gTMAX.TP_TERM
    
    
    If gIFName = "" Then
'''    If IsNumeric(txtID.Text) = False Or Len(txtID.Text) <> 5 Then
        lblErr = "* ���̵�� �н����尡 ��ġ���� �ʽ��ϴ�."
        txtPW.Text = ""
        txtID.Text = ""
        txtID.SetFocus
    Else
        lblErr = ""
'''        frmInterface.lblUserName.Caption = gIFName
'''        gTransUser = Trim(txtID.Text)
'''        gTransUserPW = Trim(txtPW.Text)
''''''        Call GetPrivateProfileString("TMAXENV", "LISUserPW", "", db_tmp, 20, App.Path & "\Interface.ini")
'''        Call WritePrivateProfileString("TMAXENV", "LISUser", gTransUser, App.Path & "\Interface.ini")
'''        Call WritePrivateProfileString("TMAXENV", "LISUserPW", gTransUserPW, App.Path & "\Interface.ini")
        gExamUID = txtID.Text
        
        frmInterface.Show 0
        Unload Me
    End If
    

End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lsWK As Integer

    If KeyCode = 13 Then
        If Len(txtID.Text) < 5 Then
            lblErr = "* ����� ���̵� �Է��ϼ���."
            txtID.SetFocus
            Exit Sub
        Else
            txtPW.SetFocus
        End If

    End If
    
End Sub

Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Len(txtPW.Text) < 3 Then
            lblErr = "* ��й�ȣ�� �Է��ϼ���."
            txtPW.SetFocus
            Exit Sub
        Else
            lblErr = ""
            lblCommit_Click
            
        End If
        
    End If
End Sub

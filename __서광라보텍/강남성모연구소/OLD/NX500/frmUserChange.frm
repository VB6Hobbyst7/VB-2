VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmUserChange 
   Caption         =   "����� ����"
   ClientHeight    =   1725
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows �⺻��
   Begin Threed.SSCommand cmdCancel 
      Height          =   345
      Left            =   1710
      TabIndex        =   6
      Top             =   1170
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   609
      _StockProps     =   78
      Caption         =   "���"
      Outline         =   0   'False
   End
   Begin Threed.SSCommand cmdCommit 
      Height          =   345
      Left            =   300
      TabIndex        =   5
      Top             =   1170
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   609
      _StockProps     =   78
      Caption         =   "Ȯ��"
      Outline         =   0   'False
   End
   Begin VB.TextBox txtPW 
      Appearance      =   0  '���
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  '��� ����
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2130
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox txtID 
      Appearance      =   0  '���
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
      Left            =   1380
      TabIndex        =   2
      Top             =   300
      Width           =   1665
   End
   Begin VB.Label lblErr 
      BackStyle       =   0  '����
      Caption         =   "* ����� ID�� Password �� �߸��Ǿ����ϴ�."
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   780
      Width           =   4635
   End
   Begin VB.Label Label2 
      Alignment       =   1  '������ ����
      Caption         =   "��й�ȣ :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   1155
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
      Left            =   150
      TabIndex        =   0
      Top             =   330
      Width           =   1125
   End
End
Attribute VB_Name = "frmUserChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblErr = ""
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCommit_Click()
Dim lsWK As Integer
    If Trim(txtID.Text) = "" Then
        lblErr = "* ����� ���̵� �Է��ϼ���."
        txtID.SetFocus
        Exit Sub
    End If
    
'    If Trim(txtPW.Text) = "" Then
'        lblErr = "* ��й�ȣ�� �Է��ϼ���."
'        txtPW.SetFocus
'        Exit Sub
'    End If
    gWorker_Info.ok = 0
    lsWK = Get_WKID(Trim(txtID.Text))
    If lsWK > 0 Then
        lblErr = ""
        frmInterface.lblUser.Caption = "����� : " & gWorker_Info.WK_NM
        Unload Me
    Else
        lblErr = "* �������� �ʴ� ���̵��Դϴ�."
        txtID.Text = ""
        txtID.SetFocus
        Exit Sub
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
'            lsWK = Get_WKID(Trim(txtID.Text))
'            If lsWK > 0 Then
                lblErr = ""
                cmdCommit_Click
'
'            Else
'                lblErr = "* �������� �ʴ� ���̵��Դϴ�."
'                txtID.Text = ""
'                txtID.SetFocus
'                Exit Sub
'            End If
        End If
    End If
    
End Sub

'Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        If Trim(txtPW.Text) = "" Then
'            lblErr = "* ��й�ȣ�� �Է��ϼ���."
'            txtPW.SetFocus
'            Exit Sub
'        Else
'            lblErr = ""
'            cmdCommit_Click
'
'        End If
'
'    End If
'End Sub


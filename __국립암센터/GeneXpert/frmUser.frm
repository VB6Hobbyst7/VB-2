VERSION 5.00
Begin VB.Form frmUser 
   Caption         =   "����� ����"
   ClientHeight    =   1230
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3855
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
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
      Left            =   1050
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1050
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "Ȯ��"
      Height          =   375
      Left            =   2610
      TabIndex        =   1
      Top             =   780
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
      Left            =   1530
      TabIndex        =   0
      Top             =   390
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
      Left            =   -270
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
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
      Left            =   210
      TabIndex        =   2
      Top             =   420
      Width           =   1125
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUser_Click()
    gIFUser = Trim(txtUser.Text)
    
    Call WritePrivateProfileString("Server", "IFUser", gIFUser, App.Path & "\interface.ini")
    
    frmInterface.lblUser = gIFUser
    
    Unload Me
End Sub

Private Sub Form_Load()
    txtUser.Text = gIFUser
    SelectFocus txtUser
End Sub


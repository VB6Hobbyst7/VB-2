VERSION 5.00
Begin VB.Form frmUser 
   Caption         =   "�����"
   ClientHeight    =   600
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdUser 
      Caption         =   "Ȯ��"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdUser_Click()
    gIFUser = Trim(txtUser.Text)
    
    Call WritePrivateProfileString("Server", "�����", gIFUser, App.Path & "\interface.ini")
    
    frmInterface.lblUser = gIFUser
    
    Unload Me
End Sub

Private Sub Form_Load()
    txtUser.Text = gIFUser
    SelectFocus txtUser
End Sub

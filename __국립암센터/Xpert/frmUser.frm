VERSION 5.00
Begin VB.Form frmUser 
   Caption         =   "사용자"
   ClientHeight    =   600
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3330
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleWidth      =   3330
   StartUpPosition =   3  'Windows 기본값
   Begin Xpert_국립암센터.MDButton cmdUser 
      Height          =   375
      Left            =   2070
      TabIndex        =   1
      Top             =   120
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "확인"
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
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
    
    Call WritePrivateProfileString("Server", "IFUser", gIFUser, App.Path & "\interface.ini")
    
    frmInterface.lblUser.Caption = gIFUser
    Unload Me
    
End Sub

Private Sub Form_Load()
    txtUser.Text = gIFUser
    
End Sub

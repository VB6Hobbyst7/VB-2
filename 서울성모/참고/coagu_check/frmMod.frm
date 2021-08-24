VERSION 5.00
Begin VB.Form frmMod 
   Caption         =   "사용자"
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdOK 
      Caption         =   "확인"
      Height          =   345
      Left            =   1230
      TabIndex        =   3
      Top             =   1050
      Width           =   855
   End
   Begin VB.TextBox txtPW 
      Height          =   300
      Left            =   870
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   870
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "PW"
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Top             =   600
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "ID"
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   210
      Width           =   405
   End
End
Attribute VB_Name = "frmMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    SQL = "select grade from user_data where userid = '" & Trim(txtUser.Text) & "' and userpw = '" & Trim(txtPW.Text) & "'"
    res = db_select_Col(gLocal, SQL)
    If Trim(gReadBuf(0)) = "0" Or Trim(gReadBuf(0)) = "1" Then
        UserState = True
        Unload Me
        
    End If
End Sub

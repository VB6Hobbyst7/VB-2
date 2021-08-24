VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   30
      ScaleHeight     =   915
      ScaleWidth      =   4605
      TabIndex        =   1
      Top             =   1740
      Width           =   4665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   1740
      TabIndex        =   0
      Top             =   750
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

 

Dim lhwnd As Long

 

lhwnd = FindWindow("notepad", vbNullString)

 

If lhwnd <> 0 Then

 

    EnumChildWindows lhwnd, AddressOf EnumChildProc, ByVal 0&

 

End If

End Sub

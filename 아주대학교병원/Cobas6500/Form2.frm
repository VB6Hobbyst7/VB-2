VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   2490
   ClientTop       =   4440
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3270
      TabIndex        =   2
      Top             =   600
      Width           =   1185
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   420
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   1245
      Left            =   630
      TabIndex        =   0
      Top             =   1680
      Width           =   3645
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Winsock1.SendData (EncodeUTF8_ADOStream(Text1))
End Sub
